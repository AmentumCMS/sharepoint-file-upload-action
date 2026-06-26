import sys
import os
import msal
import glob
import time
from office365.graph_client import GraphClient
from office365.runtime.odata.v4.upload_session_request import UploadSessionRequest
from office365.onedrive.driveitems.driveItem import DriveItem
from office365.onedrive.internal.paths.url import UrlPath
from office365.runtime.queries.upload_session import UploadSessionQuery
from office365.onedrive.driveitems.uploadable_properties import DriveItemUploadableProperties
from office365.runtime.queries.client_query import ClientQuery
from office365.runtime.queries.service_operation import ServiceOperationQuery

# office365-rest-python-client (2.5.3 through at least 2.6.2) double-slashes
# service-operation URLs (e.g. createUploadSession) on path-addressed items:
# UrlPath.segment already ends in "/" (":/name:/"), and ServiceOperationQuery.url
# joins another "/" on top, producing ".../file.iso://createUploadSession".
# Strip the trailing slash before joining to fix the URL without forking the lib.
def _fixed_service_operation_url(self):
    orig_url = ClientQuery.url.fget(self)
    if self.static:
        return "".join([self.context.service_root_url(), str(self.path)])
    return "/".join([orig_url.rstrip("/"), self.path.segment])

ServiceOperationQuery.url = property(_fixed_service_operation_url)

site_name = sys.argv[1]
sharepoint_host_name = sys.argv[2]
tenant_id = sys.argv[3]
client_id = sys.argv[4]
client_secret = sys.argv[5]
upload_path = sys.argv[6]
file_path = sys.argv[7]
max_retry = int(sys.argv[8]) or 3
login_endpoint = sys.argv[9] or "login.microsoftonline.com"
graph_endpoint = sys.argv[10] or "graph.microsoft.com"
file_path_recursive_match = sys.argv[11] if len(sys.argv) > 11 and sys.argv[11] else "False"

# below used with 'get_by_url' in GraphClient calls
tenant_url = f'https://{sharepoint_host_name}/sites/{site_name}'

# we're running this in actions, so we'll only ever have one .md file
local_files = glob.glob(file_path, recursive=file_path_recursive_match)

if not local_files:
    print(f"[Error] No files matched pattern: {file_path}")
    sys.exit(1)

def acquire_token():
    authority_url = f'https://{login_endpoint}/{tenant_id}'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id=client_id,
        client_credential=client_secret
    )
    token = app.acquire_token_for_client(scopes=[f"https://{graph_endpoint}/.default"])
    if "access_token" not in token:
        raise RuntimeError(f"Token acquisition failed: {token.get('error')}: {token.get('error_description')}")
    return token

#Replace office365 request url with the correct endpoint for non-default environments
def rewrite_endpoint(request):
    request.url = request.url.replace(
        "https://graph.microsoft.com", f"https://{graph_endpoint}"
    )

client = GraphClient(acquire_token)
client.before_execute(rewrite_endpoint, False)

def progress_status(offset, file_size):
    print(f"Uploaded {offset} bytes from {file_size} bytes ... {offset/file_size*100:.2f}%")

def success_callback(remote_file):
    print(f"[✓]File {remote_file.web_url} has been uploaded")

def resumable_upload(local_path, file_size, chunk_size, max_chunk_retry, timeout_secs):
    def _start_upload():
        with open(local_path, "rb") as local_file:
            session_request = UploadSessionRequest(
                local_file, 
                chunk_size, 
                lambda offset: progress_status(offset, file_size)
            )
            retry_seconds = timeout_secs / max_chunk_retry
            for session_request._range_data in session_request._read_next():
                for retry_number in range(max_chunk_retry):
                    try:
                        super(UploadSessionRequest, session_request).execute_query(qry)
                        break
                    except Exception as e:
                        if retry_number + 1 >= max_chunk_retry:
                            raise e
                        print(f"Retry {retry_number}: {e}")
                        time.sleep(retry_seconds)

    file_name = os.path.basename(local_path)
    full_remote_path = f"{upload_path.strip('/')}/{file_name}"
    return_type = client.sites.get_by_url(tenant_url).drive.root.get_by_path(full_remote_path)
    qry = UploadSessionQuery(
        return_type, {"item": DriveItemUploadableProperties(name=file_name)})
    return_type.context.add_query(qry).after_query_execute(_start_upload)
    return_type.get().execute_query()
    success_callback(return_type)

def upload_file(local_path, chunk_size):
    file_size = os.path.getsize(local_path)
    if file_size < chunk_size:
        drive = client.sites.get_by_url(tenant_url).drive.root.get_by_path(upload_path)
        remote_file = drive.upload_file(local_path).execute_query()
        success_callback(remote_file)
    else:
        resumable_upload(
            local_path,
            file_size,
            chunk_size,
            max_chunk_retry=60,
            timeout_secs=10*60)

for f in local_files:
  for i in range(max_retry):
    # Each retry must start from an empty query queue: execute_query() stops
    # at the first failed request and leaves everything queued after it
    # sitting in client._queries, since client/context is shared across
    # retries. Without clearing, the next retry's queries pile up behind
    # whatever didn't run yet, and the error you see can be a stale query
    # from an earlier attempt rather than the current one.
    client.clear()
    try:
        upload_file(f, 4*1024*1024)
        break
    except Exception as e:
        print(f"Unexpected error occurred: {e}, {type(e)}")
        if i == max_retry - 1:
            raise e
