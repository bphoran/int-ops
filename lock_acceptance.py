"""
Pivot weekly lock/partner hour data to paste into the Lock Acceptance.

Last Update: 2023-09-07
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

    lock = plib.aggregate_lock(FOLDER / lock_file, freq="W")
    hours = plib.weekly_hours(FOLDER / partner_file, site_map)
    partner_hours = lock.merge(
        right=hours,
        how="left",
        on=["staff_group", "partner", "shore", "week"],
    ).loc[
        :,
        [
            "partner",
            "country",
            "ecosystem",
            "channel",
            "shore",
            "staff_group",
            "week",
            "req_hours",
            "sup_hours",
        ],
    ]

    partner_hours.columns = [
        "Partner",
        "Country",
        "Ecosystem",
        "Channel",
        "Shore",
        "Staff Group",
        "week",
        "DEMAND",
        "SUPPLY",
    ]
    partner_hours.SUPPLY = partner_hours.SUPPLY.fillna(0.0)

    lock_accept = (
        partner_hours.query("Partner != 'Not Allocated'")
        .reset_index(drop=True)
        .pivot(
            index=list(partner_hours.columns[:6]),
            columns="week",
            values=["DEMAND", "SUPPLY"],
        )
        .fillna(0.0)
    )

    # Place DEMAND/SUPPLY below the week row, sort by week
    lock_accept.columns = lock_accept.columns.swaplevel(0, 1)
    lock_accept.sort_index(axis=1, inplace=True)

    plib.save_df(lock_accept, name="lock_acceptance", index=True)


if __name__ == "__main__":
    main()
