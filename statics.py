import os
from dotenv import load_dotenv

load_dotenv()


def get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


# =========================
# JOBVITE
# =========================

JOBVITE_API_KEY = get_required_env("JOBVITE_API_KEY")
JOBVITE_API_SECRET = get_required_env("JOBVITE_API_SECRET")
JOBVITE_COMPANY_ID = get_required_env("JOBVITE_COMPANY_ID")

JOBVITE_JOB_URL = get_required_env("JOBVITE_JOB_URL")
JOBVITE_CANDIDATE_URL = get_required_env("JOBVITE_CANDIDATE_URL")

JOBVITE_HEADERS = {
    "x-jvi-api": JOBVITE_API_KEY,
    "x-jvi-sc": JOBVITE_API_SECRET,
    "X-Company-Id": JOBVITE_COMPANY_ID,
    "Accept": "application/json",
    "Content-Type": "application/json",
}


# =========================
# HH.RU / HH.KZ
# =========================

HH_ACCESS_TOKEN = get_required_env("HH_ACCESS_TOKEN")
HH_USER_AGENT = get_required_env("HH_USER_AGENT")

HH_EMPLOYER_ID = get_required_env("HH_EMPLOYER_ID")
HH_CLIENT_ID = get_required_env("HH_CLIENT_ID")
HH_CLIENT_SECRET = get_required_env("HH_CLIENT_SECRET")
HH_REDIRECT_URI = get_required_env("HH_REDIRECT_URI")

HH_AUTH_URL = get_required_env("HH_AUTH_URL")
HH_TOKEN_URL = get_required_env("HH_TOKEN_URL")
HH_API_URL = get_required_env("HH_API_URL")

HH_HEADERS = {
    "Authorization": f"Bearer {HH_ACCESS_TOKEN}",
    "HH-User-Agent": HH_USER_AGENT,
    "Accept": "application/json",
}


# =========================
# NAUKRI / GULF
# =========================

NAUKRI_SECRET_KEY = get_required_env("NAUKRI_SECRET_KEY")
NAUKRI_ACCESS_KEY = get_required_env("NAUKRI_ACCESS_KEY")