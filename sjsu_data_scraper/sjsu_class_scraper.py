import re
import pandas as pd
import requests
from bs4 import BeautifulSoup

CATALOG_URL = "https://catalog.sjsu.edu/preview_program.php?catoid=13&poid=7829&returnto=4973"
SCHEDULE_URL = "https://sjsu.edu/classes/schedules/spring-2026.php"

OUTPUT_FILE = "sjsu_major_courses_spring_2026.csv"


def get_html(url: str) -> str:
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.text


def extract_major_courses_from_catalog(catalog_html: str) -> pd.DataFrame:
    soup = BeautifulSoup(catalog_html, "html.parser")
    text = soup.get_text("\n", strip=True)

    start_idx = text.find("Major Courses")
    if start_idx == -1:
        raise ValueError("Couldn't find 'Major Courses' section")

    end_markers = [
        "Approved Upper Division Electives",
        "Total Units Required",
        "Minor Requirements"
    ]

    end_idx = len(text)
    for marker in end_markers:
        idx = text.find(marker, start_idx)
        if idx != -1:
            end_idx = min(end_idx, idx)

    major_text = text[start_idx:end_idx]

    pattern = re.compile(
        r"\b([A-Z]{2,4}\s+\d+[A-Z]?)\s*-\s*(.+?)\s+\d+\s*unit",
        re.IGNORECASE
    )

    rows = []
    seen = set()

    for code, title in pattern.findall(major_text):
        code = re.sub(r"\s+", " ", code.upper()).strip()
        title = " ".join(title.split()).strip()

        if (code, title) not in seen:
            seen.add((code, title))
            rows.append({
                "course": code,
                "course title": title
            })

    return pd.DataFrame(rows)


def extract_schedule_table(schedule_html: str) -> pd.DataFrame:
    tables = pd.read_html(schedule_html)

    for df in tables:
        cols = [str(c).lower() for c in df.columns]
        if "section" in " ".join(cols) and "open seats" in " ".join(cols):
            return df

    raise ValueError("Schedule table not found")


def normalize_schedule_table(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip().lower() for c in df.columns]

    section_col = next(c for c in df.columns if "section" in c)
    title_col = next(c for c in df.columns if "course title" in c)
    seats_col = next(c for c in df.columns if "open seats" in c)

    df = df[[section_col, title_col, seats_col]].copy()
    df.columns = ["section_raw", "course title", "open seats"]

    df["course"] = df["section_raw"].str.extract(r"^([A-Z]{2,4}\s+\d+[A-Z]?)")

    df["course"] = df["course"].str.upper().str.strip()
    df["open seats"] = pd.to_numeric(df["open seats"], errors="coerce").fillna(0).astype(int)

    return df[["course", "course title", "open seats"]]


def summarize(major_df: pd.DataFrame, schedule_df: pd.DataFrame) -> pd.DataFrame:
    merged = schedule_df.merge(major_df, on="course", how="inner")

    summary = (
        merged.groupby(["course", "course title_y"], as_index=False)
        .agg(
            section_amount=("course", "size"),
            amount_of_open_seats=("open seats", "sum")
        )
        .rename(columns={"course title_y": "course title"})
    )

    final = major_df.merge(summary, on=["course", "course title"], how="left")

    final["section_amount"] = final["section_amount"].fillna(0).astype(int)
    final["amount_of_open_seats"] = final["amount_of_open_seats"].fillna(0).astype(int)

    return final[["course", "course title", "section_amount", "amount_of_open_seats"]]


def main():
    catalog_html = get_html(CATALOG_URL)
    schedule_html = get_html(SCHEDULE_URL)

    major_df = extract_major_courses_from_catalog(catalog_html)
    schedule_raw = extract_schedule_table(schedule_html)
    schedule_df = normalize_schedule_table(schedule_raw)

    result = summarize(major_df, schedule_df)

    # ✅ Save to CSV
    result.to_csv(OUTPUT_FILE, index=False)

    print("Saved to:", OUTPUT_FILE)
    print(result.head())


if __name__ == "__main__":
    main()