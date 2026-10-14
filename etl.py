import os
import sqlite3
import warnings

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(
    os.environ.get("JLI_DATA_DIR", os.path.join(BASE_DIR, "data")),
    "risk.db",
)
CONSOLIDATED_FUND_NAMES = {
    "Jubilee Life Aggressive Fund",
    "Jubilee Life Balanced Fund",
    "Meesaq Balanced Fund",
    "Aggressive Takaful Fund",
    "Balanced Takaful Fund",
    "Family Takaful Balance Fund",
}


def _normalize_symbol(value):
    symbol = str(value).strip().upper()
    if symbol.endswith(" INDEX"):
        symbol = symbol[:-6].strip()
    return symbol


def _normalize_fund_name(value):
    return " ".join(str(value).strip().split()).casefold()


def _find_sheet_name(workbook, sheet_name, path):
    actual_sheet = next(
        (
            name
            for name in workbook.sheet_names
            if name.strip().casefold() == sheet_name.casefold()
        ),
        None,
    )
    if actual_sheet is None:
        raise ValueError(f"Sheet '{sheet_name}' was not found in workbook: {path}")
    return actual_sheet


def _read_price_matrix(path, sheet_name):
    with pd.ExcelFile(path) as workbook:
        actual_sheet = _find_sheet_name(workbook, sheet_name, path)
        raw = pd.read_excel(workbook, sheet_name=actual_sheet, header=None)

    header_row = None
    date_positions = []
    for row_index in range(min(len(raw), 25)):
        date_positions = [
            column_index
            for column_index, value in enumerate(raw.iloc[row_index])
            if isinstance(value, str) and value.strip().casefold() == "date"
        ]
        if date_positions:
            header_row = row_index
            break

    if header_row is None:
        raise ValueError(
            f"No row containing a 'Date' header was found in sheet "
            f"'{actual_sheet}' of workbook: {path}"
        )

    date_position = date_positions[0]
    next_date_position = next(
        (position for position in date_positions[1:] if position > date_position),
        raw.shape[1],
    )
    header = raw.iloc[header_row]
    symbol_positions = [
        column_index
        for column_index in range(date_position + 1, next_date_position)
        if isinstance(header.iloc[column_index], str)
        and header.iloc[column_index].strip()
    ]
    symbols = [
        _normalize_symbol(header.iloc[column_index])
        for column_index in symbol_positions
    ]

    if not symbols:
        raise ValueError(
            f"No company-symbol columns were found after the Date column in "
            f"sheet '{actual_sheet}' of workbook: {path}"
        )
    if len(symbols) != len(set(symbols)):
        raise ValueError(
            f"Duplicate company-symbol headers were found in sheet "
            f"'{actual_sheet}' of workbook: {path}"
        )

    data = raw.iloc[header_row + 1 :, [date_position, *symbol_positions]].copy()
    data.columns = ["date", *symbols]
    data["date"] = pd.to_datetime(data["date"], errors="coerce").dt.normalize()
    data = data.dropna(subset=["date"])

    if data["date"].duplicated().any():
        raise ValueError(
            f"Duplicate dates were found in sheet '{actual_sheet}' "
            f"of workbook: {path}"
        )

    values = data[symbols]
    numeric_values = values.apply(pd.to_numeric, errors="coerce")
    invalid_values = values.notna() & numeric_values.isna()
    if invalid_values.to_numpy().any():
        row_offset, column_offset = invalid_values.to_numpy().nonzero()
        row_offset = row_offset[0]
        column_offset = column_offset[0]
        symbol = symbols[column_offset]
        raise ValueError(
            f"Non-numeric values were found for symbol '{symbol}' "
            f"in sheet '{actual_sheet}' of workbook: "
            f"{values.iloc[row_offset, column_offset]!r}"
        )
    data = pd.concat([data[["date"]], numeric_values], axis=1)

    return data.set_index("date")


def _read_w2_portfolio_layout(path):
    with pd.ExcelFile(path) as workbook:
        sheet_name = _find_sheet_name(workbook, "PortW", path)
        raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)

    header_row = None
    symbol_column = None
    for row_index in range(min(len(raw), 25)):
        for column_index, value in enumerate(raw.iloc[row_index]):
            if isinstance(value, str) and value.strip().casefold() == "scrips":
                header_row = row_index
                symbol_column = column_index
                break
        if header_row is not None:
            break

    if header_row is None or symbol_column is None:
        raise ValueError(f"No 'Scrips' header was found in sheet '{sheet_name}' of workbook: {path}")

    symbols = {
        _normalize_symbol(value)
        for value in raw.iloc[header_row + 1 :, symbol_column]
        if isinstance(value, str)
        and value.strip()
        and value.strip() not in {"-", "0"}
    }
    funds = {}
    for column_index in range(symbol_column + 2, raw.shape[1]):
        value = raw.iloc[header_row, column_index]
        if not isinstance(value, str) or not value.strip():
            continue
        fund_name = " ".join(value.strip().split())
        if fund_name.casefold() != "total":
            funds[_normalize_fund_name(fund_name)] = fund_name

    if not symbols:
        raise ValueError(f"No portfolio symbols were found in sheet '{sheet_name}' of workbook: {path}")
    if not funds:
        raise ValueError(f"No fund columns were found in sheet '{sheet_name}' of workbook: {path}")
    return symbols, funds


def _read_w2_security_metadata(path):
    with pd.ExcelFile(path) as workbook:
        sheet_name = _find_sheet_name(workbook, "Sectors", path)
        raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)

    header_row = None
    columns = {}
    for row_index in range(min(len(raw), 10)):
        labels = {}
        for column_index, value in enumerate(raw.iloc[row_index]):
            if isinstance(value, str) and value.strip():
                labels.setdefault(value.strip().casefold(), column_index)
        if (
            "script" in labels
            and "sector" in labels
            and "free float" in labels
        ):
            header_row = row_index
            columns = labels
            break
    if header_row is None:
        raise ValueError(f"No Script/Sector header was found in W2 sheet '{sheet_name}'.")

    script_column = columns["script"]
    sector_column = columns["sector"]
    market_cap_column = next(
        (
            index
            for index, value in enumerate(raw.iloc[header_row])
            if isinstance(value, str)
            and value.strip().casefold() == "market cap"
        ),
        None,
    )

    metadata = {}
    for row_index in range(header_row + 1, len(raw)):
        symbol_value = raw.iloc[row_index, script_column]
        if pd.isna(symbol_value) or not str(symbol_value).strip():
            continue
        symbol = _normalize_symbol(symbol_value)
        sector_value = raw.iloc[row_index, sector_column]
        free_float = raw.iloc[row_index, columns["ff %"]]
        market_cap_band = (
            raw.iloc[row_index, market_cap_column]
            if market_cap_column is not None
            else None
        )
        metadata[symbol] = {
            "sector": (
                str(sector_value).strip()
                if pd.notna(sector_value) and str(sector_value).strip()
                else None
            ),
            "free_float_pct": pd.to_numeric(pd.Series([free_float]), errors="coerce").iloc[0],
            "market_cap_band": (
                str(market_cap_band).strip()
                if pd.notna(market_cap_band)
                and str(market_cap_band).strip()
                and not str(market_cap_band).startswith("#")
                else None
            ),
        }
    return metadata


def _read_w2_security_profiles(path, target_funds, target_symbols):
    aliases = {
        "consolidated": "Consolidated",
        "jlbf": "Jubilee Life Balanced Fund",
        "jlaf": "Jubilee Life Aggressive Fund",
        "atf": "Aggressive Takaful Fund",
        "meesaq": "Meesaq Balanced Fund",
        "mtf": "Managed Takaful Fund",
        "ygf": "Yaqeen Growth Fund",
    }
    profiles = {}
    with pd.ExcelFile(path) as workbook:
        for sheet_name in workbook.sheet_names:
            sheet_key = sheet_name.strip().casefold()
            fund_name = aliases.get(sheet_key) or target_funds.get(
                _normalize_fund_name(sheet_name)
            )
            if fund_name is None:
                continue
            fund_key = _normalize_fund_name(fund_name)
            if fund_key != "consolidated" and fund_key not in target_funds:
                continue
            raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)

            header_row = None
            symbol_column = None
            for row_index in range(min(len(raw), 40)):
                labels = {
                    str(value).strip().casefold(): column_index
                    for column_index, value in enumerate(raw.iloc[row_index])
                    if isinstance(value, str) and value.strip()
                }
                if "symbol" in labels and "sector" in labels:
                    header_row = row_index
                    symbol_column = labels["symbol"]
                    break
            if header_row is None or symbol_column is None:
                continue

            measure_labels = {}
            for column_index in range(raw.shape[1]):
                label_parts = [
                    " ".join(str(raw.iloc[row_index, column_index]).strip().casefold().split())
                    for row_index in (header_row, header_row + 1)
                    if isinstance(raw.iloc[row_index, column_index], str)
                    and raw.iloc[row_index, column_index].strip()
                ]
                if label_parts:
                    measure_labels[" ".join(label_parts)] = column_index
            risk_column = next(
                (
                    column_index
                    for label, column_index in measure_labels.items()
                    if "risk" in label and "level" in label
                ),
                None,
            )
            liquidity_column = next(
                (
                    column_index
                    for label, column_index in measure_labels.items()
                    if "liquidity" in label and "basket" in label
                ),
                None,
            )

            for row_index in range(header_row + 2, len(raw)):
                symbol_value = raw.iloc[row_index, symbol_column]
                if not isinstance(symbol_value, str) or not symbol_value.strip():
                    continue
                symbol = _normalize_symbol(symbol_value)
                if symbol not in target_symbols:
                    continue
                values = {}
                for field, column_index in (
                    ("risk_level", risk_column),
                    ("liquidity_basket", liquidity_column),
                ):
                    value = (
                        raw.iloc[row_index, column_index]
                        if column_index is not None
                        else None
                    )
                    values[field] = (
                        str(value).strip()
                        if pd.notna(value) and str(value).strip()
                        else None
                    )
                if values["risk_level"] not in {"V.H.R", "H.R", "M.R", "L.R"}:
                    values["risk_level"] = None
                if values["liquidity_basket"] not in {
                    "Highly Liquid",
                    "Liquid",
                    "Semi Liquid",
                    "Illiquid",
                }:
                    values["liquidity_basket"] = None
                profiles[(fund_key, symbol)] = values
    return profiles


def _read_w2_liquidity_baskets(path):
    categories = ("Highly Liquid", "Liquid", "Semi Liquid", "Illiquid")
    records = []

    with pd.ExcelFile(path) as workbook:
        for sheet_name in workbook.sheet_names:
            raw = pd.read_excel(
                workbook, sheet_name=sheet_name, header=None, nrows=20
            )
            basket_position = None
            for row_index in range(len(raw)):
                for column_index, value in enumerate(raw.iloc[row_index]):
                    if (
                        isinstance(value, str)
                        and value.strip().casefold() == "liquidity basket"
                    ):
                        basket_position = (row_index, column_index)
                        break
                if basket_position is not None:
                    break
            if basket_position is None:
                continue

            fund_name = None
            if sheet_name.strip().casefold() == "consolidated":
                fund_name = "Consolidated"
            else:
                sheet_fund_names = {
                    "jlbf": "Jubilee Life Balanced Fund",
                    "jlaf": "Jubilee Life Aggressive Fund",
                    "atf": "Aggressive Takaful Fund",
                }
                fund_name = sheet_fund_names.get(sheet_name.strip().casefold())
                for row_index in range(min(8, len(raw))):
                    if fund_name is not None:
                        break
                    for value in raw.iloc[row_index]:
                        if (
                            isinstance(value, str)
                            and "fund" in value.strip().casefold()
                            and "risk parameters" not in value.strip().casefold()
                        ):
                            fund_name = " ".join(value.strip().split())
                            break
                    if fund_name is not None:
                        break

            if fund_name is None:
                raise ValueError(
                    f"Could not identify the fund name for the liquidity basket "
                    f"in sheet '{sheet_name}' of workbook: {path}"
                )

            row_index, column_index = basket_position
            for category in categories:
                category_value = None
                for candidate_row in range(row_index + 1, min(row_index + 7, len(raw))):
                    value = raw.iloc[candidate_row, column_index]
                    if (
                        isinstance(value, str)
                        and value.strip().casefold() == category.casefold()
                    ):
                        category_value = raw.iloc[candidate_row, column_index + 1]
                        break
                if category_value is None or pd.isna(category_value):
                    value = 0.0
                else:
                    value = pd.to_numeric(
                        pd.Series([category_value]), errors="coerce"
                    ).iloc[0]
                    if pd.isna(value):
                        raise ValueError(
                            f"Invalid liquidity value for '{fund_name}' / "
                            f"'{category}' in sheet '{sheet_name}': "
                            f"{category_value!r}"
                        )
                records.append(
                    {
                        "fund_name": fund_name,
                        "category": category,
                        "value": float(value),
                    }
                )

    if not records:
        raise ValueError(
            f"No liquidity baskets were found in the W2 workbook: {path}"
        )
    return pd.DataFrame(records).drop_duplicates(
        subset=["fund_name", "category"], keep="last"
    )


def _read_w3_share_positions(path, target_symbols, available_price_symbols, target_funds):
    with pd.ExcelFile(path) as workbook:
        sheet_name = _find_sheet_name(workbook, "Shares report", path)
        raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None)

    header_row = None
    symbol_column = None
    for row_index in range(min(len(raw), 25)):
        labels = {
            str(value).strip().casefold(): column_index
            for column_index, value in enumerate(raw.iloc[row_index])
            if isinstance(value, str)
        }
        if "fund symbol" in labels:
            header_row = row_index
            symbol_column = labels["fund symbol"]
            break

    if header_row is None or symbol_column is None:
        raise ValueError(f"No 'Fund Symbol' header was found in sheet '{sheet_name}' of workbook: {path}")

    report_date = None
    for row_index in range(min(len(raw), 5)):
        for column_index, value in enumerate(raw.iloc[row_index]):
            if isinstance(value, str) and "as on date" in value.strip().casefold():
                if column_index + 1 < raw.shape[1]:
                    report_date = pd.to_datetime(
                        raw.iloc[row_index, column_index + 1], errors="coerce"
                    )
                break
        if report_date is not None:
            break
    if report_date is None or pd.isna(report_date):
        raise ValueError(f"No valid 'As on Date' was found in sheet '{sheet_name}' of workbook: {path}")
    report_date = report_date.normalize()

    group_row = header_row + 1
    measure_row = header_row + 2
    holding_columns = []
    for column_index, value in enumerate(raw.iloc[measure_row]):
        if not isinstance(value, str) or value.strip().casefold() != "holding":
            continue
        group_name = raw.iloc[group_row, column_index]
        if not isinstance(group_name, str) or not group_name.strip():
            raise ValueError(
                f"A holding column has no fund name in sheet '{sheet_name}' "
                f"of workbook: {path}"
            )
        holding_columns.append(
            (column_index, " ".join(group_name.strip().split()))
        )

    if not holding_columns:
        raise ValueError(f"No fund holding columns were found in sheet '{sheet_name}' of workbook: {path}")

    positions = []
    skipped_symbols = set()
    missing_w1_prices = set()
    skipped_funds = set()
    for row_index in range(measure_row + 1, len(raw)):
        symbol_value = raw.iloc[row_index, symbol_column]
        if not isinstance(symbol_value, str) or not symbol_value.strip():
            continue
        symbol = _normalize_symbol(symbol_value)
        for column_index, fund_name in holding_columns:
            value = raw.iloc[row_index, column_index]
            if pd.isna(value) or (isinstance(value, str) and value.strip() in {"", "-"}):
                continue
            holding = pd.to_numeric(value, errors="coerce")
            if pd.isna(holding):
                raise ValueError(
                    f"Invalid holding for '{fund_name}' / '{symbol}' at "
                    f"{sheet_name}!{row_index + 1}: {value!r}"
                )
            if holding == 0:
                continue
            fund_key = _normalize_fund_name(fund_name)
            if fund_key not in target_funds:
                skipped_funds.add(fund_name)
            elif symbol not in target_symbols:
                skipped_symbols.add(symbol)
            elif symbol not in available_price_symbols:
                missing_w1_prices.add(symbol)
            else:
                positions.append(
                    {
                        "fund_name": target_funds[fund_key],
                        "scrip_code": symbol,
                        "holding": float(holding),
                    }
                )

    if skipped_symbols:
        warnings.warn(
            "W3 share holdings were skipped because their symbols are not in "
            f"W2 PortW: {', '.join(sorted(skipped_symbols))}",
            stacklevel=2,
        )
    if missing_w1_prices:
        warnings.warn(
            "W3 share holdings were skipped because W1 has no price history "
            f"for their symbols: {', '.join(sorted(missing_w1_prices))}",
            stacklevel=2,
        )
    if skipped_funds:
        warnings.warn(
            "W3 share holdings were skipped because their fund names are not "
            f"in W2 PortW: {', '.join(sorted(skipped_funds))}",
            stacklevel=2,
        )

    if not positions:
        raise ValueError(
            f"No W3 share holdings matched W2 PortW funds and symbols in sheet '{sheet_name}'."
        )

    holdings = pd.DataFrame(positions)
    holdings = holdings.groupby(
        ["fund_name", "scrip_code"], as_index=False
    )["holding"].sum()
    return holdings, report_date


def run_etl(w1_path, w2_path, w3_path):
    df_w1 = _read_price_matrix(w1_path, "price")
    df_volume = _read_price_matrix(w1_path, "volume")
    df_liquidity = _read_w2_liquidity_baskets(w2_path)
    target_symbols, target_funds = _read_w2_portfolio_layout(w2_path)
    security_metadata = _read_w2_security_metadata(w2_path)
    security_profiles = _read_w2_security_profiles(
        w2_path, target_funds, target_symbols
    )
    if not df_w1.index.equals(df_volume.index):
        raise ValueError("W1 Price and Volume sheets must have the same dates.")
    extra_volume_symbols = sorted(set(df_volume.columns) - set(df_w1.columns))
    if extra_volume_symbols:
        raise ValueError(
            "W1 Volume contains symbols without matching W1 Price columns: "
            f"{', '.join(extra_volume_symbols)}"
        )

    available_price_symbols = set(df_w1.columns)
    price_history_symbols = target_symbols | {"KSE30"}
    df_prices = df_w1.reindex(
        columns=[
            symbol
            for symbol in df_w1.columns
            if symbol in price_history_symbols
        ]
    )
    holdings, snapshot_date = _read_w3_share_positions(
        w3_path, target_symbols, available_price_symbols, target_funds
    )

    prices_before_snapshot = df_prices.loc[df_prices.index <= snapshot_date].ffill()
    if prices_before_snapshot.empty:
        raise ValueError(
            f"No W1 prices are available on or before W3 report date {snapshot_date.date()}."
        )
    latest_prices = prices_before_snapshot.iloc[-1]

    def add_portfolio_history(fund_name, fund_holdings):
        symbols = fund_holdings["scrip_code"].tolist()
        fund_prices = prices_before_snapshot.reindex(columns=symbols)
        values = fund_prices.mul(
            fund_holdings.set_index("scrip_code")["holding"], axis="columns"
        ).sum(axis=1, min_count=len(symbols))
        values = values.dropna()
        return pd.DataFrame(
            {
                "fund_name": fund_name,
                "date": values.index,
                "price": values.to_numpy(),
                "scrip_code": "PORTFOLIO",
            }
        )

    history_frames = [
        add_portfolio_history(fund_name, fund_holdings)
        for fund_name, fund_holdings in holdings.groupby("fund_name", sort=False)
    ]
    consolidated_source = holdings[
        holdings["fund_name"].isin(CONSOLIDATED_FUND_NAMES)
    ]
    missing_consolidated_funds = (
        CONSOLIDATED_FUND_NAMES - set(consolidated_source["fund_name"])
    )
    if missing_consolidated_funds:
        raise ValueError(
            "Cannot build the six-fund consolidated portfolio; W3 holdings "
            "are missing: "
            + ", ".join(sorted(missing_consolidated_funds))
        )
    consolidated_holdings = consolidated_source.groupby(
        "scrip_code", as_index=False
    )["holding"].sum()
    history_frames.append(
        add_portfolio_history("Consolidated", consolidated_holdings)
    )
    df_history = pd.concat(history_frames, ignore_index=True)

    df_weights = holdings.copy()
    df_weights["price"] = df_weights["scrip_code"].map(latest_prices)
    missing_prices = sorted(
        df_weights.loc[df_weights["price"].isna(), "scrip_code"].unique()
    )
    if missing_prices:
        raise ValueError(
            "No W1 price is available on or before the W3 report date for "
            f"held symbols: {', '.join(missing_prices)}"
        )
    df_weights["market_value"] = df_weights["holding"] * df_weights["price"]

    consolidated_weights = (
        df_weights.groupby("scrip_code", as_index=False)
        .agg(holding=("holding", "sum"), price=("price", "first"))
    )
    consolidated_weights["fund_name"] = "Consolidated"
    consolidated_weights["market_value"] = (
        consolidated_weights["holding"] * consolidated_weights["price"]
    )
    df_weights = pd.concat([df_weights, consolidated_weights], ignore_index=True)
    df_weights["date"] = snapshot_date
    totals = df_weights.groupby("fund_name")["market_value"].transform("sum")
    if (totals == 0).any():
        raise ValueError("Cannot calculate portfolio weights for a zero-value fund.")
    df_weights["portfolio_weight"] = df_weights["market_value"] / totals
    df_weights = df_weights[
        [
            "fund_name",
            "date",
            "scrip_code",
            "holding",
            "price",
            "market_value",
            "portfolio_weight",
        ]
    ]

    df_price_history = (
        df_prices.rename_axis("date")
        .reset_index()
        .melt(id_vars="date", var_name="scrip_code", value_name="price")
        .dropna(subset=["price"])
    )
    df_price_history = df_price_history[
        df_price_history["scrip_code"].isin(df_prices.columns)
    ]
    df_volume_history = (
        df_volume.rename_axis("date")
        .reset_index()
        .melt(id_vars="date", var_name="scrip_code", value_name="volume")
        .dropna(subset=["volume"])
    )
    df_security_metadata = pd.DataFrame.from_dict(
        security_metadata, orient="index"
    ).rename_axis("scrip_code").reset_index()
    profile_rows = [
        {
            "fund_name": target_funds.get(fund_key, "Consolidated"),
            "scrip_code": symbol,
            **profile,
        }
        for (fund_key, symbol), profile in security_profiles.items()
    ]
    df_security_profiles = pd.DataFrame(
        profile_rows,
        columns=[
            "fund_name",
            "scrip_code",
            "risk_level",
            "liquidity_basket",
        ],
    )

    conn = sqlite3.connect(DB_PATH)
    try:
        df_price_history.to_sql(
            "fact_scrip_price_history", conn, if_exists="replace", index=False
        )
        df_history.to_sql(
            "fact_portfolio_history", conn, if_exists="replace", index=False
        )
        df_weights.to_sql(
            "fact_portfolio_weights", conn, if_exists="replace", index=False
        )
        df_liquidity.to_sql(
            "fact_liquidity_basket", conn, if_exists="replace", index=False
        )
        df_volume_history.to_sql(
            "fact_scrip_volume_history", conn, if_exists="replace", index=False
        )
        df_security_metadata.to_sql(
            "dim_security_metadata", conn, if_exists="replace", index=False
        )
        df_security_profiles.to_sql(
            "fact_portfolio_security_profile",
            conn,
            if_exists="replace",
            index=False,
        )
    finally:
        conn.close()

    print(
        "ETL Success: W1 prices mapped to W2 PortfolioStc and W3 company "
        f"holdings applied ({len(df_history)} portfolio history records, "
        f"{len(df_liquidity)} liquidity basket records, "
        f"{len(df_weights)} portfolio weight records, "
        f"{len(df_volume_history)} W1 volume records)."
    )
