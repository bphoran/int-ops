"""Library of commonly used functions for manipulating partner data.

Last Update: 2023-09-12
"""

from pathlib import Path
from typing import List
from datetime import datetime
import pandas as pd
import numpy as np


FOLDER = Path(".")


def save_df(
    df: List[pd.DataFrame] | pd.DataFrame,
    sheet_names: List[str] = None,
    name: str = "test",
    to_excel: bool = False,
    path: Path = FOLDER,
    index: bool = False,
) -> None:
    if isinstance(df, pd.DataFrame):
        df = [df]

    if sheet_names is None:
        sheet_names = [f"Sheet{x+1}" for x in range(len(df))]
    elif len(sheet_names) < len(df):
        remaining_sheets = [
            f"Sheet{x+1}" for x in range(len(df)) if x > len(sheet_names) - 1
        ]
        sheet_names = sheet_names + remaining_sheets

    if to_excel:
        filename = name + ".xlsx"
        with pd.ExcelWriter(
            path / filename,
            datetime_format="m/d/yy",
        ) as writer:
            for i, tmp in enumerate(df):
                tmp.to_excel(writer, sheet_name=sheet_names[i], index=index)
    else:
        filename = name + ".csv"
        for tmp in df:
            tmp.to_csv(path / filename, index=index)

    print(f"{filename} saved.")


def build_site_map(path: Path) -> pd.DataFrame:
    """
    Load into DataFrame the location properties csv found at the provided path.
    Anaplan file location: [Model] Intuit Customer Success - PROD / Partner Portal - Production
      - SYS404 P4 Partner Location Properties
    """
    df = pd.read_csv(
        path,
        usecols=[
            1,  # Item
            # 3,  # Partner
            4,  # Partner Code
            6,  # Partner Detail
            7,  # Site Category Type
            8,  # Country
            # 9,  # Site Type
        ],
        header=0,
        names=[
            "Site",
            "Company",
            "Sector",
            "shore",
            "Country",
        ],
    )

    df["partner"] = df.Sector.str.extract(r"(.*\s+)", expand=False).str.strip()

    site_map = (
        df.loc[
            :,
            [
                "Site",
                "Sector",
                "partner",
                "Company",
                "Country",
                "shore",
            ],
        ]
        .sort_values(by=["partner", "Site"])
        .reset_index(drop=True)
    )

    return site_map


def build_staff_map(path: Path) -> pd.DataFrame:
    """
    Load into DataFrame the staff group properties csv found at the provided path.
    Anaplan file location: [Model] Intuit Customer Success - PROD / Expert Planning - Production
      - SYS352 T4 Staff Group Properties
    """
    df = pd.read_csv(
        path,
        usecols=[
            6,  # S1 Business Unit
            7,  # S15 Business Unit Segment
            8,  # T2 Eco Parent
            9,  # T3 SDT Tier
            10,  # S3 Parent Staff Group
            16,  # Channel
            17,  # Country
        ],
        header=0,
        names=[
            "BU",
            "Segment",
            "Platform",
            "Ecosystem",
            "StaffGroup",
            "Channel",
            "Country",
        ],
    )

    staff_map = (
        df.loc[
            :,
            [
                "StaffGroup",
                "Ecosystem",
                "Platform",
                "Segment",
                "BU",
                "Country",
                "Channel",
            ],
        ]
        .sort_values(by=["StaffGroup"])
        .reset_index(drop=True)
    )

    return staff_map


def aggregate_lock(path: Path, freq: str = "W") -> pd.DataFrame:
    """Load the daily lock file, aggregate prod hours at either the weekly level "W"
    or monthly level "M".


    Args:
        path (Path): Path to lock csv file.
        freq (str): "W" for weekly, "M" for monthly.


    Returns:
        pd.DataFrame: lock DataFrame.
    """
    df = pd.read_csv(
        path,
        usecols=[
            2,  # T4 Staff Group
            5,  # Vendor Site Name
            7,  # SDT Ecosystem
            8,  # Date
            17,  # Country
            18,  # Channel
            19,  # Productive Hours Calculation
        ],
        header=0,
        names=[
            "staff_group",
            "sector",
            "ecosystem",
            "date",
            "country",
            "channel",
            "req_hours",
        ],
        parse_dates=["date"],
    )

    if freq in "Dd":
        freq_label = "date"

    else:
        if freq in "Mm":
            freq_label = "month"
            df[freq_label] = df.date.mask(
                df.date.dt.day != 1,
                df.date - pd.offsets.MonthBegin(),
            )
        else:
            freq_label = "week"
            df[freq_label] = df.date.mask(
                df.date.dt.day_of_week != 6,
                df.date - pd.offsets.Week(weekday=6),
            )

        df = (
            df.groupby(
                [
                    "staff_group",
                    "sector",
                    "ecosystem",
                    freq_label,
                    "country",
                    "channel",
                ]
            )
            .sum(numeric_only=True)
            .reset_index()
        )

    df["partner"] = df.sector.str.extract(r"(.*\s+)", expand=False).str.strip()
    df["shore"] = df.sector.str.split().str[-1]

    lock = df.loc[
        :,
        [
            "country",
            "ecosystem",
            "staff_group",
            "channel",
            "partner",
            "shore",
            freq_label,
            "req_hours",
        ],
    ]

    return lock


def weekly_hours(path: Path, site_map: pd.DataFrame) -> pd.DataFrame:
    """Load the partner committed hours.


    Args:
        path (Path): Path to csv file.
        site_map (pd.DataFram): Site map DataFrame.


    Returns:
        pd.DataFrame: Weekly hours DataFrame
    """
    df = pd.read_csv(path)

    time_lables = list(df.columns[3:])
    timestamps = [
        pd.Timestamp(datetime.strptime(label[4:], "%d %b %y")) for label in time_lables
    ]

    df.columns = ["Site", "staff_group", "metric"] + timestamps
    df = df.melt(
        id_vars=["Site", "staff_group", "metric"],
        var_name="week",
        value_name="sup_hours",
    )

    weekly_hours = (
        df.loc[
            df["sup_hours"] > 0,
            [
                "staff_group",
                "Site",
                "week",
                "sup_hours",
            ],
        ]
        .reset_index(drop=True)
        .merge(
            right=site_map.loc[
                :,
                [
                    "Site",
                    "partner",
                    "shore",
                ],
            ],
            how="left",
            on="Site",
        )
        .loc[
            :,
            [
                "staff_group",
                "partner",
                "shore",
                "week",
                "sup_hours",
            ],
        ]
        .groupby(
            [
                "staff_group",
                "partner",
                "shore",
                "week",
            ]
        )
        .sum(numeric_only=True)
        .reset_index()
    )

    return weekly_hours


def daily_hours(path: Path, site_map: pd.DataFrame) -> pd.DataFrame:
    """Load the partner committed hours.


    Args:
        path (Path): Path to csv file.
        site_map (pd.DataFram): Site map DataFrame.


    Returns:
        pd.DataFrame: Daily hours DataFrame
    """
    df = pd.read_csv(path)

    time_lables = list(df.columns[3:])
    timestamps = [
        pd.Timestamp(datetime.strptime(label, "%d %b %y")) for label in time_lables
    ]

    df.columns = ["Site", "staff_group", "metric"] + timestamps
    df = df.melt(
        id_vars=["Site", "staff_group", "metric"],
        var_name="date",
        value_name="sup_hours",
    )

    daily_hours = (
        df.loc[
            df["sup_hours"] > 0,
            [
                "staff_group",
                "Site",
                "date",
                "sup_hours",
            ],
        ]
        .reset_index(drop=True)
        .merge(
            right=site_map.loc[
                :,
                [
                    "Site",
                    "partner",
                    "shore",
                ],
            ],
            how="left",
            on="Site",
        )
        .loc[
            :,
            [
                "staff_group",
                "partner",
                "shore",
                "date",
                "sup_hours",
            ],
        ]
        .groupby(
            [
                "staff_group",
                "partner",
                "shore",
                "date",
            ]
        )
        .sum(numeric_only=True)
        .reset_index()
    )

    return daily_hours
