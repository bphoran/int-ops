"""
Filter daily lock file for max prod hour day by staff group.

Last Update: 2023-09-14
"""

from pathlib import Path
import partner_library as plib


FOLDER = Path(".")


def main() -> None:
    lock_file = input("Enter daily lock file name (with .csv): ")

    lock = plib.aggregate_lock(FOLDER / lock_file, freq="D")

    sg_lock = (
        lock.groupby(["country", "ecosystem", "staff_group", "channel", "date"])
        .sum(numeric_only=True)
        .reset_index()
    )

    max_sg = sg_lock.iloc[
        sg_lock.groupby("staff_group")["req_hours"].idxmax().tolist(), :
    ]

    plib.save_df(max_sg, name="max_sg")


if __name__ == "__main__":
    main()
