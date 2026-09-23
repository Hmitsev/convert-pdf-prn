import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import zipfile
import io
import re

DATABASE_URL = "postgresql://neondb_owner:npg_WMSBX9bAzd1v@ep-flat-thunder-b59qh9dk-pooler.c-7.us-east-2.aws.neon.tech/neondb?sslmode=require&channel_binding=require"

ZIP_FILES = [
    "cross_ref_1.zip",
    "cross_ref_2.zip",
    "cross_ref_3.zip"
]

BATCH_SIZE = 10000


def normalize(value):

    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    return re.sub(
        r"[^A-Z0-9]",
        "",
        value
    )


conn = psycopg2.connect(
    DATABASE_URL
)

cur = conn.cursor()

total_rows = 0

for zip_name in ZIP_FILES:

    print(f"\nReading {zip_name}")

    with zipfile.ZipFile(zip_name, "r") as z:

        csv_name = z.namelist()[0]

        raw_bytes = z.read(csv_name)

        buffer = io.BytesIO(raw_bytes)

        df = pd.read_csv(
            buffer,
            dtype=str,
            sep=None,
            engine="python"
        )
        print(df.columns.tolist())
        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        print(
            f"Rows found: {len(df)}"
        )

        batch = []

        imported = 0

        for _, row in df.iterrows():
        if imported % 100000 == 0:
    	    print("Processing...")
            item_no = str(
                row.get(
                    "Item No.",
                    ""
                )
            ).strip()

            cross_ref = str(
                row.get(
                    "Cross-Reference No.",
                    ""
                )
            ).strip()

            vendor_no = str(
                row.get(
                    "Cross-Reference Type No.",
                    ""
                )
            ).strip()

            if cross_ref == "":
                continue

            batch.append(
                (
                    vendor_no,
                    cross_ref,
                    item_no,
                    normalize(cross_ref),
                    normalize(item_no),
                    csv_name
                )
            )

            if len(batch) >= BATCH_SIZE:

                execute_values(
                    cur,
                    """
                    INSERT INTO cross_references
                    (
                        vendor_no,
                        cross_reference_no,
                        item_no,
                        normalized_cross_reference,
                        normalized_item_no,
                        source_file
                    )
                    VALUES %s
                    ON CONFLICT DO NOTHING
                    """,
                    batch
                )

                conn.commit()

                imported += len(batch)

                print(
                    f"Imported: {imported}"
                )

                batch = []

        if batch:

            execute_values(
                cur,
                """
                INSERT INTO cross_references
                (
                    vendor_no,
                    cross_reference_no,
                    item_no,
                    normalized_cross_reference,
                    normalized_item_no,
                    source_file
                )
                VALUES %s
                ON CONFLICT DO NOTHING
                """,
                batch
            )

            conn.commit()

            imported += len(batch)

        total_rows += imported

        print(
            f"Finished {zip_name}: {imported}"
        )

cur.close()
conn.close()

print(
    f"\nTOTAL IMPORTED: {total_rows}"
)
