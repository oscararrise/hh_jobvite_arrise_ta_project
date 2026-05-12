import json
import re
import time
import random
import traceback
from typing import Any, Dict, List, Optional

import requests

from config import (
    api_url,
    headers_obtain_managers,
    get_all_managers,
    get_all_vacancies_by_managers,
    get_top_candidates,
    create_jobvite_candidate,
    save_publication_report,
    now_str,
    get_code,
    get_token
)

FORCE_REQ_ID = None


def build_hh_session(token_hhru: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token_hhru}",
        "User-Agent": "RU.hh integration (oscar.jimenez@company.com)",
        "Accept": "application/json",
    })
    return session


def safe_json(response: requests.Response) -> Any:
    try:
        return response.json()
    except Exception:
        return {"raw_text": response.text[:3000]}


def log_http_error(response: requests.Response, context: str = "") -> None:
    print("\n" + "=" * 90)
    print(f"[HTTP ERROR] {context}")
    print(f"Status code : {response.status_code}")
    print(f"Method      : {response.request.method}")
    print(f"URL         : {response.url}")

    try:
        print(json.dumps(response.json(), ensure_ascii=False, indent=2))
    except Exception:
        print(response.text[:3000])

    print("=" * 90 + "\n")


def hh_get(
    session: requests.Session,
    url: str,
    timeout: int = 30,
    params: Optional[Dict[str, Any]] = None,
    max_retries: int = 4,
) -> requests.Response:
    """
    GET con reintentos para llamadas a HH.
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, params=params, timeout=timeout)

            if response.status_code >= 400:
                log_http_error(response, context=f"HH GET failed | try={attempt}/{max_retries}")

            response.raise_for_status()
            return response

        except (
            requests.exceptions.ConnectionError,
            requests.exceptions.Timeout,
            requests.exceptions.ChunkedEncodingError,
            requests.exceptions.RequestException,
        ) as e:
            print(f"[RETRY] hh_get attempt {attempt}/{max_retries} | url={url} | error={e}")

            if attempt == max_retries:
                raise

            sleep_time = (2 ** attempt) + random.uniform(0.3, 1.2)
            print(f"[WAIT] Reconnecting in {sleep_time:.2f}s...")
            time.sleep(sleep_time)


def clean_html(html_text: Optional[str]) -> str:
    if not html_text:
        return ""
    return re.sub(r"<.*?>", "", html_text, flags=re.DOTALL).strip()


def extract_req_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    match = re.search(r"\bREQ\s*[:#-]?\s*(\d+)\b", text, flags=re.IGNORECASE)
    return match.group(1) if match else None


def get_vacancy_detail(session: requests.Session, vacancy_id: str) -> Dict[str, Any]:
    url = f"{api_url}/vacancies/{vacancy_id}"
    response = hh_get(session, url)
    return response.json()


def is_connection_error(exc: Exception) -> bool:
    text = str(exc).lower()
    connection_signals = [
        "remotedisconnected",
        "connection aborted",
        "connection reset",
        "remote end closed connection without response",
        "max retries exceeded",
        "timed out",
        "timeout",
        "temporarily unavailable",
        "protocolerror",
    ]
    return any(signal in text for signal in connection_signals)


def retry_get_top_candidates(
    token_hhru: str,
    vacancy_id: str,
    max_retries: int = 4,
) -> List[Dict[str, Any]]:
    """
    Reintenta get_top_candidates cuando hay error de red/conexión.
    """
    last_exception = None

    for attempt in range(1, max_retries + 1):
        try:
            print(f"[INFO] Getting candidates | vacancy_id={vacancy_id} | try={attempt}/{max_retries}")
            candidates = get_top_candidates(token_hhru, vacancy_id)
            return candidates or []

        except Exception as e:
            last_exception = e
            print(f"[WARN] get_top_candidates failed | vacancy_id={vacancy_id} | try={attempt}/{max_retries} | error={e}")

            if attempt == max_retries or not is_connection_error(e):
                break

            sleep_time = (2 ** attempt) + random.uniform(0.5, 1.5)
            print(f"[WAIT] Connection issue detected. Retrying in {sleep_time:.2f}s...")
            time.sleep(sleep_time)

    print(
        f"[ERROR] Could not retrieve candidates after retries | "
        f"vacancy_id={vacancy_id} | error={last_exception}"
    )
    return []


def publish_job_hhru():
    """
    Main flow:
    1. Retrieves managers from HH.
    2. Retrieves active vacancies per manager.
    3. Reads details of each vacancy.
    4. Extracts REQ from the description.
    5. Retrieves top candidates.
    6. Creates candidate in Jobvite.
    7. If Jobvite returns a real success response, saves the details in Excel.

    The Excel file is generated/updated only with successful postings.
    """
    execution_started_at = now_str()
    report_rows: List[Dict[str, Any]] = []

    stats = {
        "execution_started_at": execution_started_at,
        "vacancies_found": 0,
        "vacancies_with_req": 0,
        "vacancies_without_req": 0,
        "candidates_evaluated": 0,
        "candidates_posted_successfully": 0,
        "candidates_failed": 0,
    }

    try:
        #code = get_code()
        #token_hhru = get_token(code)

        headers_for_managers = dict(headers_obtain_managers)
        headers_for_managers.update({
            "Authorization": f"Bearer {token_hhru}",
            "User-Agent": "RU.hh integration (oscar.jimenez@company.com)",
            "Accept": "application/json",
        })

        hh_session = build_hh_session(token_hhru)

        print("[INFO] Obtaining managers...")
        managers = get_all_managers(api_url, headers_for_managers)
        print(f"[INFO] Managers found: {len(managers) if managers else 0}")

        print("[INFO] Obtaining active vacancies by managers...")
        hhru_vacancies = get_all_vacancies_by_managers(
            token=token_hhru,
            manager_ids=managers,
            base_url=api_url
        )

        stats["vacancies_found"] = len(hhru_vacancies)
        print(f"[INFO] Vacancies found: {stats['vacancies_found']}")

        for vacancy in hhru_vacancies:
            vacancy_id = str(vacancy.get("id", "")).strip()
            vacancy_name = vacancy.get("name", "")
            vacancy_url = vacancy.get("url", "")

            if not vacancy_id:
                print("[WARN] Vacancy without ID. Skipping.")
                stats["vacancies_without_req"] += 1
                continue

            try:
                print(f"\n[INFO] Processing vacancy_id={vacancy_id} | vacancy_name={vacancy_name}")

                try:
                    vacancy_payload = get_vacancy_detail(hh_session, vacancy_id)
                except Exception as e:
                    print(f"[WARN] First vacancy detail call failed | vacancy_id={vacancy_id} | error={e}")
                    print("[INFO] Rebuilding HH session and retrying vacancy detail...")
                    hh_session.close()
                    hh_session = build_hh_session(token_hhru)
                    vacancy_payload = get_vacancy_detail(hh_session, vacancy_id)

                desc = clean_html(vacancy_payload.get("description", ""))
                req = extract_req_from_text(desc)

                if FORCE_REQ_ID:
                    req = str(FORCE_REQ_ID)

                if not req:
                    stats["vacancies_without_req"] += 1
                    print(
                        f"[WARN] REQ not found in HH vacancy | "
                        f"vacancy_id={vacancy_id} | vacancy_name={vacancy_name}"
                    )
                    continue

                stats["vacancies_with_req"] += 1
                print(f"[INFO] REQ found: {req}")

                candidates = retry_get_top_candidates(token_hhru, vacancy_id, max_retries=4)
                stats["candidates_evaluated"] += len(candidates)

                print(f"[INFO] Candidates retrieved for vacancy_id={vacancy_id}: {len(candidates)}")

                for candidate in candidates:
                    try:
                        result = create_jobvite_candidate(candidate, req)

                        if result.get("success"):
                            stats["candidates_posted_successfully"] += 1
                            report_rows.append({
                                "execution_started_at": execution_started_at,
                                "posted_at": now_str(),
                                "requisition_id": str(req),
                                "hh_vacancy_id": vacancy_id,
                                "hh_vacancy_name": vacancy_name,
                                "hh_vacancy_url": vacancy_payload.get("alternate_url") or vacancy_url,
                                "candidate_first_name": candidate.get("first_name"),
                                "candidate_last_name": candidate.get("last_name"),
                                "candidate_email": candidate.get("email"),
                                "candidate_phone": candidate.get("phone"),
                                "candidate_score": candidate.get("score"),
                                "candidate_city": candidate.get("city"),
                                "candidate_resume_id": candidate.get("resume_id"),
                                "candidate_resume_url": candidate.get("alternate_resume_url") or candidate.get("resume_url"),
                                "candidate_negotiation_id": candidate.get("negotiation_id"),
                                "jobvite_status_code": result.get("status_code"),
                                "jobvite_success": "YES",
                                "jobvite_message": result.get("message"),
                                "jobvite_response_id": result.get("response_id"),
                            })

                            print(
                                f"[OK] Candidate created successfully: "
                                f"{candidate.get('first_name', '')} {candidate.get('last_name', '')} "
                                f"| req={req} | vacancy_id={vacancy_id}"
                            )
                        else:
                            stats["candidates_failed"] += 1
                            print(
                                f"[ERROR] Candidate not created: "
                                f"{candidate.get('first_name', '')} {candidate.get('last_name', '')} "
                                f"| req={req} | detail={result.get('message')}"
                            )

                        time.sleep(0.15)

                    except Exception as e:
                        stats["candidates_failed"] += 1
                        print(
                            f"[ERROR] Error posting candidate in vacancy_id={vacancy_id} "
                            f"| req={req} | error={e}\n{traceback.format_exc()}"
                        )

                print(
                    f"[INFO] Vacancy processed | vacancy_id={vacancy_id} | "
                    f"vacancy_name={vacancy_name} | req={req} | candidates={len(candidates)}"
                )

            except Exception as e:
                print(
                    f"[ERROR] Error processing vacancy {vacancy_id} | "
                    f"error={e}\n{traceback.format_exc()}"
                )

    except Exception as e:
        print(
            "[FATAL] Error in publish_job_hhru:",
            str(e),
            traceback.format_exc()
        )

    finally:
        try:
            report_info = save_publication_report(report_rows, stats)

            print("\n================== FINAL SUMMARY ==================")
            print(f"Execution started at: {stats['execution_started_at']}")
            print(f"Vacancies found: {stats['vacancies_found']}")
            print(f"Vacancies with REQ: {stats['vacancies_with_req']}")
            print(f"Vacancies without REQ: {stats['vacancies_without_req']}")
            print(f"Candidates evaluated: {stats['candidates_evaluated']}")
            print(f"Candidates successfully posted: {stats['candidates_posted_successfully']}")
            print(f"Candidates failed: {stats['candidates_failed']}")
            print(f"Excel file: {report_info.get('file_path')}")
            print("===================================================\n")

        except Exception as e:
            print(f"[ERROR] Failed to save Excel report: {e}")


if __name__ == "__main__":
    publish_job_hhru()