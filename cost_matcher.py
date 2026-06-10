"""
토더 메뉴별 판매량 × 식부자재 단가표 → 메뉴별 식자재/부자재 원가 집계

매칭 단계:
  1) 정확 매칭 (메뉴명 동일)
  2) 별칭 매핑 (ALIAS: 토더명 → 식부자재명)
  3) 주류/음료 포괄명 → 대표/평균 단가
  4) 제외 목록 (배달비, 할인, 추가메뉴, 리뷰/서비스 0원 항목 등)
  5) 세트/콤보: 이름을 '+' '&'로 분해해 구성품 단가 합산
  6) 그래도 미매칭 → 별도 목록(수기 보완용)
"""
import re

# ── 1+1·단품 별칭 (토더명 → 식부자재명) ─────────────────────
ALIAS = {
    "두툼감자채빠삭전":   "두툼감자채빠삭파전",
    "해물오꼬노미빠삭전": "해물오꼬노미빠삭파전",
    "치즈콘마요빠삭전":   "치즈콘마요빠삭파전",
    "즉석두부김치":       "즉석두부돼지김치",
}

# ── 주류/음료 포괄명 → 대표 단가 키(식부자재명) 또는 평균그룹 ──
#   값이 str이면 해당 식부자재명 단가 사용,
#   값이 list면 해당 식부자재명들의 평균 단가 사용
DRINK_MAP = {
    "소주":               ["진로", "참이슬", "처음처럼", "새로"],   # 평균
    "맥주":               ["카스", "테라", "한맥"],
    "콜라":               "코카콜라",
    "제로 콜라":          "코카콜라제로",
    "제로 사이다":        "사이다",
    "국순당 바밤바밤":    "바밤바밤막걸리",
    "제주우도 땅콩 막걸리":"우도땅콩막걸리",
    "가덕,장수 등 기본막걸리": ["칠구막걸리", "호랑이생막걸리"],
    "토닉워터 300ml":     "토닉워터",
}

# ── 원가 0 처리 / 집계 제외 (비메뉴·정산성 항목) ───────────
#   정규식: 이 패턴에 맞으면 원가 0 (매출엔 포함되나 식부자재 원가는 없음)
EXCLUDE_PATTERNS = [
    r"배달비", r"할인", r"쿠폰", r"주문금액 할인",
    r"\(리뷰\)", r"\(서비스\)", r"리뷰 이벤트", r"본사 권장",
    r"추가$", r"추가 ", r"사리", r"소스", r"청양고추", r"설탕",
    r"기타$", r"기타 ", r"얼음컵", r"조리", r"필요 없음",
    r"O/X", r"빼고", r"^기본 숯불맛", r"숯불맛$",
]


# 토더 세트에서 쓰이는 약어 → 식부자재 정식명
ABBREV = {
    "빠파": "빠삭파전",
}


def _expand_abbrev(s: str) -> str:
    for a, full in ABBREV.items():
        s = s.replace(a, full)
    return s


# ── 세트 BOM (구성품 정의) ──────────────────────────────
#   구성품은 식부자재 메뉴명 또는 그룹토큰(GROUP_TOKENS).
#   '택1' 선택지는 해당 메뉴들의 평균 단가로 처리.
SET_BOM = {
    "친구야 한잔 세트":  ["@파전택1", "비빔골뱅이&소면"],
    "깊은 밤 SET":       ["@파전택1", "비빔골뱅이&소면", "@찌개택1"],
    "가벼운 밤 SET":     ["@파전택1", "바삭고추튀김", "비빔골뱅이&소면"],
    "시그니처세트 2":    ["@파전택1", "비빔골뱅이&소면"],
    "내 맘대로 파전 1+1":["@파전택1"],   # 1+1 = 세트 수량 기준 파전 1개
}

# 그룹토큰 → '택1' 평균 대상 메뉴명 목록 (또는 동적 카테고리 매칭)
GROUP_TOKENS = {
    "@찌개택1": ["순두부찌개", "바지락조개탕"],
    # @파전택1 은 '빠삭파전/빠삭전' 포함 메뉴 전체 평균 (동적)
}


def _norm(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", "", s)
    return s


def build_cost_lookup(sik_bu_rows: list[dict]) -> dict:
    """
    식부자재 행 리스트 → {정규화메뉴명: {'name','sik','bu'}}
    sik_bu_rows: [{'name','sik','bu'}, ...]
    """
    lut = {}
    for row in sik_bu_rows:
        name = row["name"]
        if not name:
            continue
        lut[_norm(name)] = {
            "name": name,
            "sik":  float(row["sik"] or 0),
            "bu":   float(row["bu"] or 0),
        }
    return lut


def _drink_cost(toorder_name: str, lut: dict):
    spec = DRINK_MAP.get(toorder_name.strip())
    if spec is None:
        return None
    if isinstance(spec, str):
        hit = lut.get(_norm(spec))
        return (hit["sik"], hit["bu"]) if hit else None
    # 평균
    vals = [lut[_norm(s)] for s in spec if _norm(s) in lut]
    if not vals:
        return None
    return (sum(v["sik"] for v in vals)/len(vals),
            sum(v["bu"] for v in vals)/len(vals))


def _is_excluded(name: str) -> bool:
    return any(re.search(p, name) for p in EXCLUDE_PATTERNS)


def _match_one(name: str, lut: dict):
    """단일 메뉴명 → (sik, bu, 매칭방식) 또는 None"""
    key = _norm(name)
    # 1) 정확
    if key in lut:
        return lut[key]["sik"], lut[key]["bu"], "정확"
    # 2) 별칭
    if name.strip() in ALIAS:
        ak = _norm(ALIAS[name.strip()])
        if ak in lut:
            return lut[ak]["sik"], lut[ak]["bu"], "별칭"
    # 3) 주류/음료 포괄명
    d = _drink_cost(name, lut)
    if d:
        return d[0], d[1], "주류포괄"
    return None


def _resolve_component(spec: str, lut: dict):
    """BOM 구성품 1개 → (sik, bu). 그룹토큰/택1 평균 처리. 실패 시 None."""
    spec = spec.strip()
    # 그룹토큰
    if spec == "@파전택1":
        items = [v for k, v in lut.items() if "빠삭파전" in k or "빠삭전" in k]
        if not items:
            return None
        return (sum(i["sik"] for i in items)/len(items),
                sum(i["bu"] for i in items)/len(items))
    if spec in GROUP_TOKENS:
        names = GROUP_TOKENS[spec]
        items = [lut[_norm(n)] for n in names if _norm(n) in lut]
        if not items:
            return None
        return (sum(i["sik"] for i in items)/len(items),
                sum(i["bu"] for i in items)/len(items))
    # 일반 메뉴명
    m = _match_one(spec, lut)
    return (m[0], m[1]) if m else None


def _match_bom(name: str, lut: dict):
    """SET_BOM에 정의된 세트 → 구성품 단가 합산."""
    bom = SET_BOM.get(name.strip())
    if not bom:
        return None
    sik = bu = 0.0
    hit = 0
    for comp in bom:
        r = _resolve_component(comp, lut)
        if r:
            sik += r[0]; bu += r[1]; hit += 1
    if hit == 0:
        return None
    return sik, bu, f"세트BOM({hit}/{len(bom)})"


def _match_set(name: str, lut: dict):
    """세트/콤보: 이름을 '+' '&'로 분해해 구성품 단가 합산."""
    # 대괄호 설명 제거 후 구분자 분해
    base = re.sub(r"\[.*?\]", "", name)
    base = re.sub(r"\(.*?\)", "", base)
    base = base.replace("세트", "").replace("SET", "").replace("1+1", "")
    parts = re.split(r"[+&]", base)
    parts = [_expand_abbrev(p.strip()) for p in parts if p.strip()]
    if len(parts) < 2:
        return None
    sik = bu = 0.0
    hit = 0
    for p in parts:
        m = _match_one(p, lut)
        if m:
            sik += m[0]; bu += m[1]; hit += 1
    if hit == 0:
        return None
    return sik, bu, f"세트분해({hit}/{len(parts)})"


def match_costs(toorder_rows: list[dict], sik_bu_rows: list[dict]) -> dict:
    """
    toorder_rows: [{'category','name','qty','sales'}, ...]
    sik_bu_rows : [{'name','sik','bu'}, ...]
    Returns dict:
      total_sik, total_bu, total_sales,
      matched_sales, unmatched_sales,
      rows(list with per-menu detail), unmatched(list)
    """
    lut = build_cost_lookup(sik_bu_rows)
    total_sik = total_bu = total_sales = matched_sales = 0.0
    rows, unmatched = [], []

    for tr in toorder_rows:
        name = str(tr["name"]).strip()
        qty  = float(tr.get("qty") or 0)
        sales= float(tr.get("sales") or 0)
        total_sales += sales

        if _is_excluded(name):
            rows.append({**tr, "sik_unit":0,"bu_unit":0,"sik":0,"bu":0,"method":"제외"})
            continue

        m = _match_one(name, lut) or _match_bom(name, lut) or _match_set(name, lut)
        if m:
            sik_u, bu_u, method = m
            sik, bu = sik_u*qty, bu_u*qty
            total_sik += sik; total_bu += bu; matched_sales += sales
            rows.append({**tr, "sik_unit":sik_u,"bu_unit":bu_u,
                         "sik":sik,"bu":bu,"method":method})
        else:
            unmatched.append(tr)
            rows.append({**tr, "sik_unit":None,"bu_unit":None,
                         "sik":0,"bu":0,"method":"미매칭"})

    return {
        "total_sik": total_sik, "total_bu": total_bu,
        "total_sales": total_sales,
        "matched_sales": matched_sales,
        "unmatched_sales": total_sales - matched_sales,
        "rows": rows, "unmatched": unmatched,
    }
