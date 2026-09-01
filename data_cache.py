import datetime
import os
import pickle
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from zoneinfo import ZoneInfo

import pandas as pd


class FeishuAPIError(RuntimeError):
    def __init__(self, message: str, code: Optional[int] = None):
        super().__init__(message)
        self.code = code


class MonthlyQuotaExceeded(FeishuAPIError):
    pass


def parse_feishu_response(response, action: str) -> dict:
    try:
        payload = response.json()
    except ValueError as error:
        raise FeishuAPIError(f"{action}失败：飞书返回了非 JSON 响应") from error

    code = payload.get("code", 0)
    if response.status_code >= 400 or code != 0:
        message = payload.get("msg") or f"HTTP {response.status_code}"
        error_class = MonthlyQuotaExceeded if code == 99991403 else FeishuAPIError
        raise error_class(f"{action}失败：{message}（错误码 {code}）", code=code)
    return payload


@dataclass
class DataLoadResult:
    dataframe: pd.DataFrame
    source: str
    updated_at: Optional[str]
    error: Optional[Exception] = None


def load_snapshot(cache_path: Path) -> tuple[pd.DataFrame, Optional[str]]:
    if not cache_path.exists():
        return pd.DataFrame(), None
    try:
        with cache_path.open("rb") as cache_file:
            payload = pickle.load(cache_file)
        dataframe = payload.get("dataframe")
        if not isinstance(dataframe, pd.DataFrame) or dataframe.empty:
            return pd.DataFrame(), None
        return dataframe, payload.get("updated_at")
    except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
        return pd.DataFrame(), None


def save_snapshot(cache_path: Path, dataframe: pd.DataFrame, updated_at: str) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = cache_path.with_suffix(cache_path.suffix + ".tmp")
    with temp_path.open("wb") as cache_file:
        pickle.dump(
            {"dataframe": dataframe, "updated_at": updated_at},
            cache_file,
            protocol=pickle.HIGHEST_PROTOCOL,
        )
    os.replace(temp_path, cache_path)


def load_with_fallback(
    fetch_data: Callable[[], pd.DataFrame],
    cache_path: Path,
    memory_dataframe: Optional[pd.DataFrame] = None,
    memory_updated_at: Optional[str] = None,
) -> DataLoadResult:
    try:
        dataframe = fetch_data()
        if dataframe.empty:
            raise FeishuAPIError("飞书表格没有返回数据")
        updated_at = datetime.datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(
            timespec="seconds"
        )
        snapshot_error = None
        try:
            save_snapshot(cache_path, dataframe, updated_at)
        except OSError as error:
            # 部署平台若暂时不允许写本地文件，仍保留本次实时数据和内存缓存。
            snapshot_error = error
        return DataLoadResult(dataframe, "live", updated_at, snapshot_error)
    except Exception as error:
        if isinstance(memory_dataframe, pd.DataFrame) and not memory_dataframe.empty:
            return DataLoadResult(
                memory_dataframe.copy(), "cache", memory_updated_at, error
            )
        cached_dataframe, cached_at = load_snapshot(cache_path)
        if not cached_dataframe.empty:
            return DataLoadResult(cached_dataframe, "cache", cached_at, error)
        raise
