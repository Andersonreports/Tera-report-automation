from supabase import create_client
import os

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

_client = None

def _get_client():
    global _client
    if _client is None:
        if not SUPABASE_URL or not SUPABASE_KEY:
            raise RuntimeError(
                "SUPABASE_URL and SUPABASE_KEY environment variables are not set. "
                "Create a .env file or export them before starting the server."
            )
        _client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _client

# Keep 'supabase' as a module-level accessor for backward compatibility
class _LazyClient:
    def __getattr__(self, name):
        return getattr(_get_client(), name)

supabase = _LazyClient()


def upload_pdf(file_path, file_name):
    client = _get_client()
    with open(file_path, "rb") as f:
        client.storage.from_("reports").upload(
            file_name,
            f,
            {"upsert": "true", "content-type": "application/pdf"}
        )
    url = client.storage.from_("reports").get_public_url(file_name)
    return url


def save_report(user_id, file_url, report_type):
    data = {
        "user_id": user_id,
        "file_url": file_url,
        "report_type": report_type
    }
    _get_client().table("reports").insert(data).execute()