# -*- coding: utf-8 -*-
"""
中文地名脏数据清洗脚本
功能：
  1) 按逗号/顿号/分号/空格/加号/划线/竖线及「及」「和」「与」拆分多地名
  2) 格式清洗 → 去重
  3) cpca 标准化（含裸区县后缀回退）
  4) 粘连地名检测（如「常州苏州杭州」→ 处理失败，交人工）
  5) 分离「不合规 / 处理失败」问题数据

输出 result.xlsx 六列：
  源数据 | 拆分片段 | 处理后数据 | 剩余问题数据 | 问题类型 | 应删除
"""
# =========================
# 依赖安装命令（如环境已装可跳过）
# =========================
# pip install pandas openpyxl tqdm cpca

import re
import sys
import warnings
from pathlib import Path

import pandas as pd
import cpca
from tqdm import tqdm

# =========================
# 可配置路径（按需修改）
# =========================
INPUT_FILE = r"./examples/dirty_data.xlsx"   # 输入 Excel 文件路径
OUTPUT_FILE = r"./examples/result.xlsx"      # 输出 Excel 文件路径
# 读取工作表：默认第一个表；如需指定可改为表名字符串，如 "Sheet1"
SHEET_NAME = 0
# cpca 分批转换大小（便于显示进度条；数据量不大时可改大）
CPCA_BATCH_SIZE = 200
# 粘连检测：去掉已识别省/市后，残留汉字数 ≥ 该阈值时进入二次判定
RESIDUAL_THRESHOLD = 2

# 拆分分隔符：中英文逗号、顿号、中英文分号、半角/全角空格、+ / \ |
# 以及连接词「及」「和」「与」（注意：含「和」的地名如「和田/和县」会被误拆，见 README 注意事项）
_SPLIT_PATTERN = re.compile(r"[，.,、；;： 　+/\\|及和与-]+")
# 段首尾剥离的轻量标点（句号等不作分隔符，只去掉）
_EDGE_PUNCT_PATTERN = re.compile(
    r"^[\s　。．.、，,；;：:！!？?]+|[\s　。．.、，,；;：:！!？?]+$"
)
# 「等 / 等等」→ 整段不合规（不因「等」拆分）
_ETC_PATTERN = re.compile(r"等等?")
# 统计残留时使用的汉字
_CJK_PATTERN = re.compile(r"[一-鿿]")

# 行政区划后缀（长的在前，便于最长匹配）
_ADMIN_SUFFIXES = (
    "特别行政区",
    "自治州",
    "地区",
    "盟",
    "市",
    "区",
    "县",
    "旗",
    "州",
    "省",
)

# 裸区县名识别失败时的回退后缀
# 注意：不把「县」放进回退列表——cpca 对「xxx县」常返回 市='县' 的脏匹配
_FALLBACK_SUFFIXES = ("区", "市", "自治州", "地区", "盟", "州", "旗")

# 问题类型常量
ISSUE_INVALID = "不合规"       # POI、带「等」、完全无法识别等
ISSUE_FAILED = "处理失败"     # 疑似多地名粘连，交人工拆分


def clean_text(value) -> str:
    """
    格式清洗（对拆分后的单段）：
    - 去除前后空格（含全角空格）
    - 删除所有阿拉伯数字（半角 0-9 + 全角 ０-９）
    - 英文字母统一转为大写
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    text = str(value)
    text = text.strip().strip("　").strip()
    text = re.sub(r"[0-9０-９]", "", text)
    text = text.upper()
    text = text.strip().strip("　").strip()
    return text


def normalize_source(raw) -> str:
    """源数据：只做首尾空白清理，保留原句便于追溯。"""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    return str(raw).strip().strip("　").strip()


def split_raw_to_segments(raw_text: str) -> list:
    """
    按审批规则拆分多地名：
    分隔符：， , 、 ； ; 空格 全角空格 + / \\ | 及「及」「和」「与」
    不因「等」拆分；段首尾去掉句号等轻量标点。
    """
    if not raw_text:
        return []

    parts = _SPLIT_PATTERN.split(raw_text)
    segments = []
    for part in parts:
        seg = part.strip().strip("　").strip()
        # 反复去掉首尾句号/逗号等（如「广东。」→「广东」）
        while seg:
            new_seg = _EDGE_PUNCT_PATTERN.sub("", seg)
            if new_seg == seg:
                break
            seg = new_seg.strip().strip("　").strip()
        if seg:
            segments.append(seg)
    return segments


def has_etc_marker(text: str) -> bool:
    """段内出现「等」或「等等」→ 不合规。"""
    return bool(_ETC_PATTERN.search(text))


def is_valid_region(value) -> bool:
    """判断 cpca 返回的省/市/区字段是否有效。"""
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text:
        return False
    if text.lower() in {"nan", "none", "null"}:
        return False
    return True


def format_region(province, city):
    """
    组合规则：
    a) 识别出市 → 「xx省xx市」（省缺失则只输出市）
    b) 只识别出省 → 「xx省」
    c) 都没有 → None
    过滤：市字段为「县/市/区…」单独出现时整行作废。
    """
    has_province = is_valid_region(province)
    has_city = is_valid_region(city)

    if has_city:
        city_text = str(city).strip()
        if city_text in {"县", "市", "区", "旗", "州", "盟", "地区"}:
            return None
        province_text = str(province).strip() if has_province else ""
        if province_text:
            return f"{province_text}{city_text}"
        return city_text

    if has_province:
        return str(province).strip()

    return None


def strip_one_admin_suffix(name: str) -> str:
    """
    只剥一层行政区划后缀。
    注意：不可循环剥——「常州市」→「常州」即可，
    若继续剥「州」会变成「常」，破坏粘连检测。
    """
    if not name:
        return ""
    for suffix in _ADMIN_SUFFIXES:
        # _ADMIN_SUFFIXES 已按长度大致从长到短排列
        if name.endswith(suffix) and len(name) > len(suffix):
            return name[: -len(suffix)]
    return name


def chinese_only(text: str) -> str:
    """提取文本中的汉字，保持顺序。"""
    if not text:
        return ""
    return "".join(_CJK_PATTERN.findall(str(text)))


def extract_residual_after_province_city(original: str, province, city) -> str:
    """
    从原文中去掉已识别的省、市（全称 + 剥一层后缀的裸名），
    返回残留汉字串。

    故意不使用「区/地址」字段做删除：
    cpca 会把「常州苏州杭州」里未消费的「苏州杭州」放进地址，
    若把地址也删掉，残留会变成 0，粘连检测失效。
    """
    text = original or ""
    tokens = []
    for part in (province, city):
        if not is_valid_region(part):
            continue
        full = str(part).strip()
        if not full:
            continue
        tokens.append(full)
        bare = strip_one_admin_suffix(full)
        if bare and bare != full:
            tokens.append(bare)

    for token in sorted(set(tokens), key=len, reverse=True):
        if token:
            text = text.replace(token, "")

    return chinese_only(text)


def cpca_transform(location_list):
    """
    调用 cpca.transform 批量转换。
    优先 open_warning=False；cpca 0.5.5 不支持时自动回退。
    """
    if not location_list:
        return pd.DataFrame(columns=["省", "市", "区", "地址", "adcode"])

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return cpca.transform(location_list, open_warning=False)
        except TypeError:
            return cpca.transform(location_list)


def _ends_with_admin_suffix(name: str) -> bool:
    return any(name.endswith(suffix) for suffix in _ADMIN_SUFFIXES)


def parse_region_from_row(row, province_col, city_col):
    return format_region(row[province_col], row[city_col])


def _get_cols(cpca_df):
    """解析 cpca 结果列名。"""
    col_map = {str(c).strip(): c for c in cpca_df.columns}
    province_col = col_map.get("省")
    city_col = col_map.get("市")
    district_col = col_map.get("区")
    address_col = col_map.get("地址")
    if province_col is None or city_col is None:
        raise RuntimeError(
            f"cpca 返回列异常，当前列名：{list(cpca_df.columns)}，期望包含「省」「市」"
        )
    return province_col, city_col, district_col, address_col


def _batch_cpca(names, batch_size):
    """分批调用 cpca，返回对齐后的 DataFrame。"""
    if not names:
        return pd.DataFrame(columns=["省", "市", "区", "地址", "adcode"])
    frames = []
    for start in range(0, len(names), batch_size):
        frames.append(cpca_transform(names[start:start + batch_size]))
    return pd.concat(frames, ignore_index=True)


def _try_resolve_one_name(name: str, batch_size=50):
    """
    尝试解析单个地名：原文 + 后缀回退。
    返回 (processed_str_or_None, province, city, district, address)
    """
    candidates = [name]
    if name and not _ends_with_admin_suffix(name):
        candidates.extend(name + s for s in _FALLBACK_SUFFIXES)

    df = _batch_cpca(candidates, batch_size)
    province_col, city_col, district_col, address_col = _get_cols(df)

    for i in range(len(candidates)):
        row = df.iloc[i]
        processed = parse_region_from_row(row, province_col, city_col)
        if processed:
            district = row[district_col] if district_col is not None else None
            address = row[address_col] if address_col is not None else None
            return (
                processed,
                row[province_col],
                row[city_col],
                district,
                address,
            )
    return None, None, None, None, None


def is_glue_or_partial(original, province, city, district=None) -> bool:
    """
    判断是否「只匹配了局部 / 多地名粘连」。

    规则（残留阈值 = RESIDUAL_THRESHOLD）：
    1) 从原文去掉省、市后，残留汉字数 < 阈值 → 非粘连（完整匹配）
    2) 残留可被 cpca 解析：
       - 解析结果与当前省+市相同 → 视为同城下属区县信息（如「成都双流」）→ 非粘连
       - 解析结果是另一个市/省 → 粘连（如「常州苏州杭州」「上海杭州」）
    3) 残留无法解析且长度 ≥ 阈值 → 粘连/处理失败
       （如「眉山青白江及什邡」去掉眉山后仍很长）
    """
    residual = extract_residual_after_province_city(original, province, city)
    if len(residual) < RESIDUAL_THRESHOLD:
        return False

    # 残留再解析：是否另一个行政区
    re_processed, _, re_city, _, _ = _try_resolve_one_name(residual)
    if re_processed:
        parent = format_region(province, city)
        # 同一省+市（区县细节）→ 允许成功
        if parent and re_processed == parent:
            return False
        # 市名不同 → 粘连（常州 + 苏州）
        if is_valid_region(city) and is_valid_region(re_city):
            if str(city).strip() != str(re_city).strip():
                return True
            return False
        # 解析出了不同结果 → 处理失败
        if parent and re_processed != parent:
            return True

    # 残留对不上同城区县，且够长 → 处理失败
    # 额外：若区字段裸名恰好等于残留，视为区县细节
    if is_valid_region(district):
        dist_full = str(district).strip()
        dist_bare = strip_one_admin_suffix(dist_full)
        if residual == chinese_only(dist_full) or residual == chinese_only(dist_bare):
            return False

    return True


# 单字行政区划后缀，用于粘连解析时清理消费点残留的孤立后缀字
_SINGLE_SUFFIX_CHARS = set("市县区省州盟旗")


def greedy_split_locations(name: str, batch_size=50):
    """
    贪心粘连解析：从左往右逐个消费最左可解析地名。

    思路：对剩余串，从最短可能地名（2 字）递增尝试，取「能解析成合法省/市的
    最短前缀」作为本轮消费对象，记录结果后从剩余串删除，继续处理剩下的部分。

    返回 dict：
      {"ok": True,  "consumed": [(片段, 标准化结果), ...], "remaining": ""}
      {"ok": False, "consumed": [...已消费...], "remaining": 剩余无法解析的文本}
    全部消费完 → ok=True；中途解不动 → ok=False（remaining 为解不动的剩余）。
    """
    remaining = name
    consumed = []  # [(segment, processed), ...]
    for _ in range(len(name) // 2 + 1):
        if not remaining:
            return {"ok": True, "consumed": consumed, "remaining": remaining}
        found = False
        for L in range(2, len(remaining) + 1):
            prefix = remaining[:L]
            processed, _, _, _, _ = _try_resolve_one_name(prefix, batch_size)
            if processed:
                consumed.append((prefix, processed))
                remaining = remaining[L:]
                # 清理消费点后残留的孤立后缀字，如「常州市苏州市」消费「常州」后剩「市苏州市」
                while remaining and remaining[0] in _SINGLE_SUFFIX_CHARS:
                    remaining = remaining[1:]
                found = True
                break
        if not found:
            # 剩余部分解不动 → 失败，返回已消费信息供上层判断「应删除」
            return {"ok": False, "consumed": consumed, "remaining": remaining}
    return {"ok": True, "consumed": consumed, "remaining": remaining}


def _handle_glue(name: str) -> list:
    """
    对粘连段：先尝试贪心拆，能全消费则多行成功，否则处理失败。
    处理失败时，若「首个有效城市之后的剩余文本长度 ≥ 2」，标注应删除=是。
    """
    result = greedy_split_locations(name)
    if result["ok"] and result["consumed"]:
        return [
            {"kind": "ok", "segment": seg, "processed": processed}
            for seg, processed in result["consumed"]
        ]
    # 粘连拆不动 → 处理失败
    should_delete = bool(result["consumed"]) and len(result["remaining"]) >= 2
    return [
        {
            "kind": "issue",
            "issue": ISSUE_FAILED,
            "problem": name,
            "delete": should_delete,
        }
    ]


def resolve_locations(unique_cleaned, batch_size=200):
    """
    cpca 标准化 + 裸区县后缀回退 + 粘连贪心解析。

    返回 list[list[dict] | None]：
      成功: [{"kind": "ok", "processed": "四川省成都市"}, ...]
      粘连拆开: [{"kind": "ok", ...}, {"kind": "ok", ...}]（多行）
      粘连拆不动: [{"kind": "issue", "issue": "处理失败", "problem": 原文}]
      失败: None  （外层再标「不合规」）
    """
    total = len(unique_cleaned)
    batch_size = max(1, int(batch_size))

    # ---- 第一轮：原文批量转换 ----
    print("正在调用 cpca 进行批量地址解析...")
    batch_frames = []
    for start in tqdm(range(0, total, batch_size), desc="cpca转换", unit="批"):
        batch = unique_cleaned[start:start + batch_size]
        batch_frames.append(cpca_transform(batch))

    cpca_df = pd.concat(batch_frames, ignore_index=True) if batch_frames else pd.DataFrame()
    if len(cpca_df) != total:
        raise RuntimeError(
            f"cpca 返回行数与输入不一致：输入 {total}，返回 {len(cpca_df)}"
        )

    province_col, city_col, district_col, *_ = _get_cols(cpca_df)

    results = [None] * total
    pending_indices = []

    for i in range(total):
        row = cpca_df.iloc[i]
        processed = parse_region_from_row(row, province_col, city_col)
        name = unique_cleaned[i]
        district = row[district_col] if district_col is not None else None

        if processed:
            if is_glue_or_partial(
                name, row[province_col], row[city_col], district
            ):
                # 粘连：尝试贪心拆多行，拆不动才处理失败
                results[i] = _handle_glue(name)
            else:
                results[i] = [
                    {"kind": "ok", "segment": name, "processed": processed}
                ]
        else:
            if name and not _ends_with_admin_suffix(name):
                pending_indices.append(i)

    # ---- 后缀回退（双流 → 双流区）----
    if pending_indices:
        print(
            f"原文未识别 {len(pending_indices)} 条，尝试追加「区/市」等后缀回退解析..."
        )
        still_pending = pending_indices
        for suffix in _FALLBACK_SUFFIXES:
            if not still_pending:
                break

            queries = [unique_cleaned[i] + suffix for i in still_pending]
            retry_df = _batch_cpca(queries, batch_size)

            next_pending = []
            for local_idx, global_idx in enumerate(still_pending):
                row = retry_df.iloc[local_idx]
                processed = parse_region_from_row(row, province_col, city_col)
                name = unique_cleaned[global_idx]
                if processed:
                    district = row[district_col] if district_col is not None else None
                    # 残留检测仍基于「清洗后原文」，不是带后缀的查询串
                    if is_glue_or_partial(
                        name, row[province_col], row[city_col], district
                    ):
                        results[global_idx] = _handle_glue(name)
                    else:
                        results[global_idx] = [
                            {"kind": "ok", "segment": name, "processed": processed}
                        ]
                else:
                    next_pending.append(global_idx)
            still_pending = next_pending

    return results


def load_first_column(file_path, sheet_name=0):
    """读取 Excel 第一列原始数据，并做基础异常检查。"""
    path = Path(file_path)

    if not path.exists():
        print(f"[错误] 输入文件不存在：{path.resolve()}")
        print("请检查脚本顶部 INPUT_FILE 路径是否正确。")
        sys.exit(1)

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception as exc:
        print(f"[错误] 读取 Excel 失败：{exc}")
        sys.exit(1)

    if df is None or df.empty or df.shape[1] < 1:
        print("[错误] Excel 内容为空，或没有可用的第一列数据。")
        sys.exit(1)

    first_col = df.iloc[:, 0].tolist()
    if len(first_col) == 0:
        print("[错误] 第一列为空，无法处理。")
        sys.exit(1)

    return first_col


def expand_raw_rows(raw_values):
    """
    将每一行源数据拆成多段分析对象。
    返回 [(源数据整句, 清洗后片段), ...]
    源数据始终写整句；同一源可对应多行结果。
    """
    pairs = []
    for raw in raw_values:
        source = normalize_source(raw)
        if not source:
            cleaned = clean_text(raw)
            if cleaned:
                pairs.append((source, cleaned))
            continue

        segments = split_raw_to_segments(source)
        if not segments:
            segments = [source]

        for seg in segments:
            cleaned = clean_text(seg)
            if cleaned:
                pairs.append((source, cleaned))

    return pairs


def main():
    print("=" * 50)
    print("中文地名脏数据清洗开始")
    print("=" * 50)

    # ---------- 读取 ----------
    raw_values = load_first_column(INPUT_FILE, SHEET_NAME)
    print(f"已读取原始行数：{len(raw_values)}")

    # ---------- 拆分 + 格式清洗 ----------
    cleaned_pairs = expand_raw_rows(raw_values)
    if not cleaned_pairs:
        print("[错误] 清洗后没有任何有效数据（全为空或仅含数字/空白）。")
        sys.exit(1)

    print(f"拆分后片段数：{len(cleaned_pairs)}")

    # ---------- 基于清洗后片段去重（保留首次出现的源整句） ----------
    dedup_map = {}  # cleaned -> source
    for source, cleaned in cleaned_pairs:
        if cleaned not in dedup_map:
            dedup_map[cleaned] = source

    unique_cleaned = list(dedup_map.keys())
    print(f"去重后分析对象数：{len(unique_cleaned)}")

    # ---------- 预检：「等/等等」直接不合规，不送 cpca ----------
    pre_issue = {}  # cleaned -> issue type
    to_resolve = []
    for cleaned in unique_cleaned:
        if has_etc_marker(cleaned):
            pre_issue[cleaned] = ISSUE_INVALID
        else:
            to_resolve.append(cleaned)

    # ---------- cpca 标准化（含粘连检测） ----------
    resolve_map = {}  # cleaned -> result dict or None
    if to_resolve:
        try:
            resolved_list = resolve_locations(to_resolve, batch_size=CPCA_BATCH_SIZE)
        except Exception as exc:
            print(f"[错误] cpca 转换失败：{exc}")
            print("请确认已正确安装 cpca：pip install cpca")
            sys.exit(1)
        for cleaned, result in zip(to_resolve, resolved_list):
            resolve_map[cleaned] = result

    # ---------- 组装五列结果 ----------
    success_rows = []
    problem_rows = []

    for cleaned in unique_cleaned:
        source = dedup_map[cleaned]

        # 预检不合规（含「等」）
        if cleaned in pre_issue:
            problem_rows.append(
                {
                    "源数据": source,
                    "拆分片段": cleaned,
                    "处理后数据": "",
                    "剩余问题数据": cleaned,
                    "问题类型": pre_issue[cleaned],
                    "应删除": "",
                }
            )
            continue

        result = resolve_map.get(cleaned)

        if result:
            # result 为 list：可能是单个成功，或多个成功（粘连拆开），或问题
            for item in result:
                if item.get("kind") == "ok" and item.get("processed"):
                    success_rows.append(
                        {
                            "源数据": source,
                            "拆分片段": item.get("segment") or cleaned,
                            "处理后数据": item["processed"],
                            "剩余问题数据": "",
                            "问题类型": "",
                            "应删除": "",
                            "_sort_key": item["processed"],
                        }
                    )
                elif item.get("issue") == ISSUE_FAILED:
                    # 粘连：处理失败，交人工
                    problem_rows.append(
                        {
                            "源数据": source,
                            "拆分片段": cleaned,
                            "处理后数据": "",
                            "剩余问题数据": item.get("problem") or cleaned,
                            "问题类型": ISSUE_FAILED,
                            "应删除": "是" if item.get("delete") else "",
                        }
                    )
                else:
                    # 完全无法识别：不合规
                    problem_rows.append(
                        {
                            "源数据": source,
                            "拆分片段": cleaned,
                            "处理后数据": "",
                            "剩余问题数据": cleaned,
                            "问题类型": ISSUE_INVALID,
                            "应删除": "",
                        }
                    )
        else:
            # 完全无法识别：不合规
            problem_rows.append(
                {
                    "源数据": source,
                    "拆分片段": cleaned,
                    "处理后数据": "",
                    "剩余问题数据": cleaned,
                    "问题类型": ISSUE_INVALID,
                    "应删除": "",
                }
            )

    # 成功在前（按省+市排序），问题在后
    success_rows.sort(key=lambda x: x["_sort_key"])
    problem_rows.sort(
        key=lambda x: (0 if x["问题类型"] == ISSUE_INVALID else 1, x["剩余问题数据"])
    )

    result_records = []
    for row in success_rows:
        result_records.append(
            {
                "源数据": row["源数据"],
                "拆分片段": row["拆分片段"],
                "处理后数据": row["处理后数据"],
                "剩余问题数据": row["剩余问题数据"],
                "问题类型": row["问题类型"],
                "应删除": row["应删除"],
            }
        )
    for row in problem_rows:
        result_records.append(row)

    result_df = pd.DataFrame(
        result_records,
        columns=["源数据", "拆分片段", "处理后数据", "剩余问题数据", "问题类型", "应删除"],
    )

    # ---------- 写出 ----------
    try:
        result_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    except Exception as exc:
        print(f"[错误] 写出结果失败：{exc}")
        sys.exit(1)

    # ---------- 统计 ----------
    total_out = len(result_df)
    success_count = len(success_rows)
    invalid_count = sum(1 for r in problem_rows if r["问题类型"] == ISSUE_INVALID)
    failed_count = sum(1 for r in problem_rows if r["问题类型"] == ISSUE_FAILED)

    print("-" * 50)
    print(f"处理完成！结果已保存至：{Path(OUTPUT_FILE).resolve()}")
    print(f"总数：{total_out}")
    print(f"成功数：{success_count}")
    print(f"不合规数：{invalid_count}")
    print(f"处理失败数：{failed_count}")
    print(f"剩余问题合计：{invalid_count + failed_count}")
    print("=" * 50)


if __name__ == "__main__":
    main()