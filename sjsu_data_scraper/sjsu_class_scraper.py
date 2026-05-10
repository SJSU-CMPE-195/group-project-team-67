import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
from io import StringIO

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
                "course_title": title
            })

    return pd.DataFrame(rows)

def extract_schedule_table(schedule_html: str) -> pd.DataFrame:
    tables = pd.read_html(StringIO(schedule_html), flavor="lxml")
    return tables[0]

def normalize_days(days: str) -> str:
    days = days.upper().strip()

    # Handle combined patterns first
    if days == "TR":
        return "TTh"
    if days == "R":
        return "Th"

    # Replace standalone R with Th
    days = days.replace("R", "Th")

    return days

def normalize_schedule_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    section_col = df.columns[0]
    title_col = df.columns[3]
    days_col = df.columns[7]
    time_col = df.columns[8]

    df = df[[section_col, title_col, days_col, time_col]].copy()
    df.columns = ["section_raw", "course_title", "days", "time"]

    df["section_raw"] = df["section_raw"].astype(str)
    df["course"] = df["section_raw"].str.extract(r"^([A-Z]{2,4}\s+\d+[A-Z]?)")
    df["course"] = df["course"].str.upper().str.strip()

    df["days"] = df["days"].astype(str).apply(normalize_days)
    df["time"] = df["time"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    df["class_time"] = df["days"] + " " + df["time"]
    df["class_time"] = df["class_time"].str.replace("TBA TBA", "TBA").str.strip()

    return df[["course", "course_title", "class_time"]]


def summarize(major_df: pd.DataFrame, schedule_df: pd.DataFrame) -> pd.DataFrame:
    merged = schedule_df.merge(major_df, on="course", how="inner")

    summary = (
        merged.groupby(["course", "course_title_y"], as_index=False)
        .agg(
            section_amount=("course", "size"),
            valid_class_times=("class_time", lambda x: ", ".join(sorted(set(t for t in x if t and t != "nan"))))
        )
        .rename(columns={"course_title_y": "course_title"})
    )

    final = major_df.merge(summary, on=["course", "course_title"], how="left")

    final["section_amount"] = final["section_amount"].fillna(0).astype(int)
    final["valid_class_times"] = final["valid_class_times"].fillna("")

    return final[["course", "course_title", "section_amount", "valid_class_times"]]

def main():
    try:
        print("Starting script...")

        catalog_html = get_html(CATALOG_URL)
        schedule_html = get_html(SCHEDULE_URL)
        print("Fetched HTML successfully")

        major_df = extract_major_courses_from_catalog(catalog_html)
        print("Extracted major courses")

        schedule_raw = extract_schedule_table(schedule_html)
        print("Extracted schedule table")

        schedule_df = normalize_schedule_table(schedule_raw)
        print("Normalized schedule")

        result = summarize(major_df, schedule_df)

        print("About to save:", OUTPUT_FILE)
        print(result.shape)

        result.to_csv(OUTPUT_FILE, index=False)

        print("Saved to:", OUTPUT_FILE)

    except Exception as e:
        print("ERROR:", type(e).__name__)
        print(str(e)[:1000])

if __name__ == "__main__":
    main()