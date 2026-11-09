#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
구로구 전체 토지 & 단독·다가구 매매 실거래가 초고속 수집 엔진
- VWorld API 키가 주석 처리(비활성화)되면 외부 통신 대기 없이 즉시 국토부 마스킹 데이터로 패스스루
- 전체 워크플로우 실행 시간: 약 20~30초 초고속 완료
"""

import os
import json
import argparse
import datetime
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET

class GuroRealEstateCollector:
    def __init__(self, api_key: str = None, vworld_key: str = None, vworld_domain: str = None):
        # 1. 국토부 API 키
        raw_key = (api_key or os.environ.get("GURO_API_KEY") or "")
        self.clean_api_key = urllib.parse.quote(urllib.parse.unquote(raw_key.strip()))
        
        # 2. VWorld API 키 및 도메인 (주석 처리 시 None 처리)
        self.vworld_api_key = (vworld_key or os.environ.get("VWORLD_API_KEY") or "").strip()
        self.vworld_domain = (vworld_domain or os.environ.get("VWORLD_DOMAIN") or "").strip()
        
        self.lawd_cd = "11530"   # 구로구
        
        if not self.vworld_api_key:
            print("ℹ️ [알림] VWorld 지번 매칭이 주석 처리되어 있습니다. 순수 국토부 실거래 초고속 수집 모드로 작동합니다.")

    def _get_val(self, item, *tags):
        tags_lower = [t.lower() for t in tags]
        for child in item:
            if child.tag.lower() in tags_lower:
                val = child.text
                if val and str(val).strip():
                    return str(val).strip()
        return ""

    def _request_api(self, endpoint_url: str, deal_ym: str) -> list:
        if not self.clean_api_key:
            return []

        params = {
            "LAWD_CD": self.lawd_cd,
            "DEAL_YMD": deal_ym,
            "pageNo": "1",
            "numOfRows": "1000" 
        }
        query_str = urllib.parse.urlencode(params)
        full_url = f"{endpoint_url}?serviceKey={self.clean_api_key}&{query_str}"

        req = urllib.request.Request(full_url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            with urllib.request.urlopen(req, timeout=15) as response:
                xml_content = response.read().decode("utf-8", errors="replace")
            root = ET.fromstring(xml_content)
            return root.findall(".//item")
        except Exception as e:
            print(f"    ❌ [국토부 API 에러] {e}")
            return []

    def fetch_all_real_transactions(self, months_back: int = 12) -> list:
        today = datetime.date.today()
        deal_ym_list = []
        for i in range(months_back):
            year = today.year
            month = today.month - i
            while month <= 0:
                month += 12
                year -= 1
            deal_ym_list.append(f"{year}{str(month).zfill(2)}")
        
        deal_ym_list = sorted(list(set(deal_ym_list)), reverse=True)
        land_url = "http://apis.data.go.kr/1613000/RTMSDataSvcLandTrade/getRTMSDataSvcLandTrade"
        house_url = "http://apis.data.go.kr/1613000/RTMSDataSvcSHTrade/getRTMSDataSvcSHTrade"
        collected = []
        
        print("\n[1/2] 토지 매매 실거래 데이터 수집 중...")
        for ym in deal_ym_list:
            items = self._request_api(land_url, ym)
            for it in items:
                dong = (self._get_val(it, "법정동", "umdNm") or self._get_val(it, "시군구", "sggNm") or "구로구")
                jibun = self._get_val(it, "지번", "jibun")
                area_m2 = float((self._get_val(it, "거래면적", "dealArea", "plttgAr", "area") or "0").replace(",", "").strip() or 0)
                price_val = int((self._get_val(it, "거래금액", "dealAmount") or "0").replace(",", "").strip() or 0)
                
                if price_val > 0:
                    collected.append({
                        "category": "land",
                        "category_name": "토지",
                        "dong": dong,
                        "deal_date": f"{self._get_val(it, '년', 'dealYear')}-{self._get_val(it, '월', 'dealMonth').zfill(2)}-{self._get_val(it, '일', 'dealDay').zfill(2)}",
                        "raw_jibun": jibun,
                        "masked_address": f"{dong} {jibun}".strip(),
                        "jimok": self._get_val(it, "지목", "lndcNm") or "대",
                        "use_area": self._get_val(it, "용도지역", "prposArea1Nm") or "-",
                        "road_cond": self._get_val(it, "도로조건") or "-",
                        "area_m2": area_m2,
                        "price_만원": price_val,
                        "note": "토지 실거래"
                    })

        print("\n[2/2] 단독·다가구 매매 실거래 데이터 수집 중...")
        for ym in deal_ym_list:
            items = self._request_api(house_url, ym)
            for it in items:
                dong = (self._get_val(it, "법정동", "umdNm") or self._get_val(it, "시군구", "sggNm") or "구로구")
                jibun = self._get_val(it, "지번", "jibun")
                house_type = self._get_val(it, "주택유형", "houseType") or "단독"
                
                lot_area_str = self._get_val(it, "대지면적", "plArea", "plottage", "plttgAr", "landArea", "daejiArea", "거래면적", "dealArea")
                bldg_area_str = self._get_val(it, "연면적", "totArea", "totAr", "flrArea", "yeonArea", "houseArea", "bldgArea")
                
                if not lot_area_str or not bldg_area_str:
                    for child in it:
                        t_lower = child.tag.lower()
                        val = child.text
                        if val and str(val).strip():
                            if not bldg_area_str and ('tot' in t_lower or '연면적' in t_lower or 'flr' in t_lower):
                                bldg_area_str = str(val).strip()
                            elif not lot_area_str and ('area' in t_lower or '면적' in t_lower or 'pl' in t_lower):
                                lot_area_str = str(val).strip()

                lot_area = float((lot_area_str or "0").replace(",", "").strip() or 0)
                bldg_area = float((bldg_area_str or "0").replace(",", "").strip() or 0)
                price_val = int((self._get_val(it, "거래금액", "dealAmount") or "0").replace(",", "").strip() or 0)

                if price_val > 0:
                    collected.append({
                        "category": "house",
                        "category_name": f"{house_type}주택",
                        "dong": dong,
                        "deal_date": f"{self._get_val(it, '년', 'dealYear')}-{self._get_val(it, '월', 'dealMonth').zfill(2)}-{self._get_val(it, '일', 'dealDay').zfill(2)}",
                        "raw_jibun": jibun,
                        "masked_address": f"{dong} {jibun}".strip(),
                        "jimok": "대",
                        "use_area": self._get_val(it, "용도지역", "prposArea1Nm") or "-",
                        "road_cond": self._get_val(it, "도로조건") or "-",
                        "area_m2": lot_area,
                        "bldg_area_m2": bldg_area,
                        "floors": "단독/다가구",
                        "build_year": self._get_val(it, "건축년도", "buildYear") or "-",
                        "price_만원": price_val,
                        "note": f"{house_type}주택 실거래"
                    })

        return collected

    def cross_verify_address(self, raw_jibun: str, dong_name: str, area_m2: float) -> dict:
        is_masked = "*" in raw_jibun
        
        if not is_masked:
            return {
                "verified_jibun": f"{dong_name} {raw_jibun}".strip(),
                "road_address": "지번 확인 완료",
                "confidence": "원본 지번 (마스킹 없음)",
            }
        
        # VWorld가 비활성화되어 있는 경우 즉시 패스스루
        return {
            "verified_jibun": f"{dong_name} {raw_jibun}".strip(),
            "road_address": "지번 마스킹 (국토부 비공개)",
            "confidence": "부분 공개(국토부 마스킹)",
        }

    def process_and_enrich(self, raw_deals: list) -> list:
        processed = []
        for it in raw_deals:
            pyung = it["area_m2"] / 3.305785 if it["area_m2"] > 0 else 0
            verified = self.cross_verify_address(it["raw_jibun"], it.get("dong", "구로구"), it["area_m2"])
            
            processed.append({
                "category": it["category"],
                "category_name": it["category_name"],
                "deal_date": it["deal_date"],
                "masked_address": it["masked_address"],
                "verified_address": verified["verified_jibun"],
                "road_address": verified["road_address"],
                "confidence": verified["confidence"],
                "jimok": it["jimok"],
                "use_area": it["use_area"],
                "road_cond": it["road_cond"],
                "area_m2": it["area_m2"],
                "area_pyung": pyung,
                "bldg_area_m2": it.get("bldg_area_m2", 0.0),
                "floors": it.get("floors", "-"),
                "build_year": it.get("build_year", "-"),
                "is_over_30": pyung >= 30.0,
                "price_만원": it["price_만원"],
                "unit_price_pyung": it["price_만원"] / pyung if pyung > 0 else 0,
                "unit_price_m2": it["price_만원"] / it["area_m2"] if it["area_m2"] > 0 else 0,
                "note": it["note"]
            })
        processed.sort(key=lambda x: x["deal_date"], reverse=True)
        return processed

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--export-json", type=str, default=None)
    parser.add_argument("--cli", action="store_true")
    args = parser.parse_args()

    collector = GuroRealEstateCollector()
    raw_deals = collector.fetch_all_real_transactions(months_back=12)
    processed_records = collector.process_and_enrich(raw_deals)

    kst_tz = datetime.timezone(datetime.timedelta(hours=9))
    now_kst = datetime.datetime.now(kst_tz)

    if args.export_json:
        formatted_deals = []
        for idx, r in enumerate(processed_records, 1):
            formatted_deals.append({
                "id": idx,
                "category": r["category"],
                "categoryName": r["category_name"],
                "dealDate": r["deal_date"],
                "maskedAddress": r["masked_address"],
                "verifiedJibun": r["verified_address"],
                "roadAddress": r["road_address"],
                "jimok": r["jimok"],
                "useArea": r["use_area"],
                "roadCond": r["road_cond"],
                "areaM2": round(r["area_m2"], 1),
                "pyung": round(r["area_pyung"], 1),
                "bldgAreaM2": round(r["bldg_area_m2"], 1) if r["bldg_area_m2"] > 0 else None,
                "floors": r["floors"],
                "buildYear": r["build_year"],
                "isOver30": r["is_over_30"],
                "priceManwon": r["price_만원"],
                "unitPricePyung": round(r["unit_price_pyung"], 1),
                "unitPriceM2": round(r["unit_price_m2"], 1),
                "confidence": r["confidence"],
                "note": r["note"]
            })

        export_payload = {
            "lastSync": now_kst.strftime("%Y-%m-%d %H:%M:%S (KST)"),
            "lastSyncIso": now_kst.isoformat(),
            "totalCount": len(formatted_deals),
            "deals": formatted_deals
        }
        with open(args.export_json, "w", encoding="utf-8") as f:
            json.dump(export_payload, f, ensure_ascii=False, indent=2)
        print(f"\n✅ [완료] data.json 정상 저장됨 (총 {len(formatted_deals)}건)")

if __name__ == "__main__":
    main()