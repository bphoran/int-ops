"""
Merge daily lock prod hours with daily partner committed hours. Make sure the lock
csv and committed hour csv are in the same folder as this python script. User will be
prompted for the lock file name and partner hour file name, including ".csv".

Also required - SYS404 P4 Partner Location Properties.csv
Found at:
[Model] Intuit Customer Success - PROD / Partner Portal - Production
    - SYS404 P4 Partner Location Properties

Last Update: 2023-09-12
"""

from pathlib import Path
import partner_library as plib

FOLDER = Path(".")


def main() -> None:
    site_file = FOLDER / "SYS404 P4 Partner Location Properties.csv"
    if not site_file.is_file():
        print('"SYS404 P4 Partner Location Properties.csv" missing')
        return

    site_map = plib.build_site_map(site_file)
    lock_file = input("Enter daily lock file name (with .csv): ")
    partner_file = input("Enter partner hour file name (with .csv): ")

    lock = plib.aggregate_lock(FOLDER / lock_file, freq="D")
    hours = plib.daily_hours(FOLDER / partner_file, site_map)
    partner_hours = lock.merge(
        right=hours,
        how="left",
        on=[
            "staff_group",
            "partner",
            "shore",
            "date",
        ],
    )

    partner_hours.sup_hours = partner_hours.sup_hours.fillna(0.0)
    missing_supply = (
        partner_hours.query("sup_hours <= 0 and partner != 'Not Allocated'")
        .sort_values(
            [
                "date",
                "partner",
                "staff_group",
                "shore",
            ]
        )
        .reset_index(drop=True)
    )

    plib.save_df(partner_hours, name="partner_hours")
    plib.save_df(missing_supply, name="missing_supply")


if __name__ == "__main__":
    main()
