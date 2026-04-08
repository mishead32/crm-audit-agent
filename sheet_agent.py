import anyio
import gspread
import pandas as pd
from io import StringIO
from claude_agent_sdk import query, ClaudeAgentOptions, ResultMessage

# ── CONFIG ───────────────────────────────────────────────────────────────────
CREDENTIALS_PATH = r"D:\Coding Works\Visual Studio Code\BIPS IRRELEVANT AUDIT\credentials.json.json"
SHEET_URL        = "https://docs.google.com/spreadsheets/d/16uZ_W0EVzxIsBUEzVch5UjD01E_Wzhj5D9tOTD61VnY/edit"
WORKSHEET_NAME   = "Inquiry"
# ─────────────────────────────────────────────────────────────────────────────


def load_sheet_data() -> pd.DataFrame:
    client    = gspread.service_account(filename=CREDENTIALS_PATH)
    sheet     = client.open_by_url(SHEET_URL)
    worksheet = sheet.worksheet(WORKSHEET_NAME)
    rows      = worksheet.get_all_values()
    headers   = rows[0]
    seen, unique_headers = {}, []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)
    return pd.DataFrame(rows[1:], columns=unique_headers)


def build_context(df: pd.DataFrame) -> str:
    buf = StringIO()
    df.info(buf=buf)

    cat_summary = ""
    for col in df.select_dtypes(include="object").columns:
        top = df[col].value_counts().head(10)
        cat_summary += f"\n{col}:\n{top.to_string()}\n"

    return f"""=== DATASET OVERVIEW ===
Rows: {len(df)}  |  Columns: {len(df.columns)}
Columns: {', '.join(df.columns.tolist())}

=== COLUMN TYPES & NULL COUNTS ===
{buf.getvalue()}

=== NUMERIC SUMMARY ===
{df.describe(include='number').to_string() if not df.select_dtypes(include='number').empty else 'No numeric columns'}

=== CATEGORY VALUE COUNTS (top 10 each) ===
{cat_summary}

=== FULL DATA (CSV) ===
{df.to_csv(index=False)}"""


async def ask(question: str, data_context: str) -> str:
    prompt = f"""You are a data analyst. Here is a Google Sheet dataset (worksheet: "{WORKSHEET_NAME}"):

{data_context}

User question: {question}

Answer based solely on the data above. Show exact numbers and percentages where relevant."""

    result = ""
    async for message in query(
        prompt=prompt,
        options=ClaudeAgentOptions(
            allowed_tools=[],   # no file/shell tools needed — data is in the prompt
            max_turns=3,
        ),
    ):
        if isinstance(message, ResultMessage):
            result = message.result
    return result


async def main():
    print("Connecting to Google Sheets...")
    try:
        df = load_sheet_data()
    except Exception as e:
        print(f"ERROR loading sheet: {e}")
        return

    print(f"Loaded {len(df)} rows x {len(df.columns)} columns from '{WORKSHEET_NAME}'.\n")
    data_context = build_context(df)

    print("─" * 60)
    print("  Google Sheet AI Agent  —  ask anything about your data")
    print("  Type 'refresh' to reload the sheet | 'quit' to exit")
    print("─" * 60 + "\n")

    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye!")
            break

        if not user_input:
            continue
        if user_input.lower() in ("quit", "exit", "q"):
            print("Goodbye!")
            break
        if user_input.lower() == "refresh":
            print("Refreshing data from Google Sheets...")
            try:
                df = load_sheet_data()
                data_context = build_context(df)
                print(f"Refreshed: {len(df)} rows x {len(df.columns)} columns.\n")
            except Exception as e:
                print(f"ERROR refreshing: {e}")
            continue

        print("Agent: thinking...", end="\r")
        try:
            answer = await ask(user_input, data_context)
            print(" " * 20, end="\r")  # clear "thinking..." line
            print(f"Agent: {answer}\n")
        except Exception as e:
            print(f"ERROR: {e}\n")


if __name__ == "__main__":
    anyio.run(main)
