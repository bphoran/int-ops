"""
Generate partner allocations at either the weekly or monthly level.

Last Update: 2023-09-12
"""

from pathlib import Path
import partner_library as plib


FOLDER = Path(".")


def main() -> None:
    lock_file = input("Enter daily lock file name (with .csv): ")
    freq = input("For weekly allocations, type 'W', monthly type 'M': ")

    lock = plib.aggregate_lock(FOLDER / lock_file, freq=freq)

    freq_label = "month" if freq in ["M", "m"] else "week"

    partner_allocation = lock.assign(
        sg_hour_sum=lock.groupby(["staff_group", freq_label])["req_hours"].transform(
            lambda x: x.sum()
        ),
        allocation=lambda x: x.eval("req_hours / sg_hour_sum"),
    )

    plib.save_df(partner_allocation, name=f"partner_{freq_label}ly_alloc")


if __name__ == "__main__":
    main()
