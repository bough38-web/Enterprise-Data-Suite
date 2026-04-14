import pandas as pd
import numpy as np
import re
from thefuzz import process, fuzz

class DataEngine:
    @staticmethod
    def normalize_col(x):
        return str(x).replace(" ", "").replace("_", "").strip()

    @staticmethod
    def auto_find_key(df1, df2):
        priority = ["계약번호", "서비스(소)", "고객번호", "ID"]
        for col in priority:
            if col in df1.columns and col in df2.columns:
                return col
        common = [c for c in df1.columns if c in df2.columns]
        if common: return common[0]
        if not df1.columns.empty: return df1.columns[0]
        return None

    @staticmethod
    def auto_match_columns(df1, df2):
        left_map = {DataEngine.normalize_col(c): c for c in df1.columns}
        right_map = {DataEngine.normalize_col(c): c for c in df2.columns}
        
        matches = {}
        for norm, real_right in right_map.items():
            if norm in left_map:
                matches[real_right] = left_map[norm]
        return matches

    @staticmethod
    def perform_matching(df_left, df_right, key_col, match_map):
        """
        Optimized vectorized matching using pd.merge.
        """
        # Create a Copy to avoid modifying original
        df_res = df_left.copy()
        
        # Prepare Reference data for merge
        # Filter right to only include key and matched columns
        right_cols = [key_col] + list(match_map.keys())
        df_right_sub = df_right[right_cols].copy()
        
        # Ensure string keys and rename matched columns to avoid collisions
        df_right_sub[key_col] = df_right_sub[key_col].astype(str).str.strip()
        df_res[key_col] = df_res[key_col].astype(str).str.strip()
        
        rename_map = {old: f"매칭_{old}" for old in match_map.keys()}
        df_right_sub.rename(columns=rename_map, inplace=True)
        
        # Vectorized Merge (Left Join)
        # dropping duplicates on key for many-to-one matching safety
        df_right_sub.drop_duplicates(subset=[key_col], inplace=True)
        df_res = pd.merge(df_res, df_right_sub, on=key_col, how='left')
        
        return df_res

    @staticmethod
    def apply_filters(df, filter_config):
        """
        filter_config handles:
        - auto_target (시설구분/요금구분 == '대상')
        - custom_filters: list of {'column': str, 'values': list, 'mode': 'include'|'exclude'}
        Returns: (df_result, diagnostics_dict)
        """
        res = df.copy()
        diag = {
            "initial": len(res),
            "auto_target_removed": 0,
            "custom_filter_removed": 0
        }
        
        # 1. Auto Target Filtering
        if filter_config.get('auto_target'):
            pre_count = len(res)
            for col in ["시설구분", "요금구분"]:
                if col in res.columns:
                    # Case-insensitive and strip whitespace/padding
                    res = res[res[col].astype(str).str.strip() == "대상"]
            diag["auto_target_removed"] = pre_count - len(res)
                    
        # 2. Custom Filters
        if filter_config.get('custom_filters'):
            pre_count = len(res)
            for f in filter_config.get('custom_filters', []):
                col = f['column']
                vals = f['values']
                mode = f['mode']
                
                if col in res.columns and vals:
                    series = res[col].astype(str).str.strip()
                    if mode == 'include':
                        res = res[series.isin(vals)]
                    else:
                        res = res[~series.isin(vals)]
            diag["custom_filter_removed"] = pre_count - len(res)
                    
        diag["final"] = len(res)
        return res, diag

    @staticmethod
    def select_columns(df, columns, mode='keep'):
        if mode == 'keep':
            valid_cols = [c for c in columns if c in df.columns]
            return df[valid_cols] if valid_cols else df
        else:
            return df.drop(columns=[c for c in columns if c in df.columns], errors='ignore')

    @staticmethod
    def apply_replacements(df, replacements):
        """
        replacements: list of dicts like [{'column': 'ColA', 'find': 'old', 'replace': 'new', 'exact': False}]
        """
        res = df.copy()
        for rule in replacements:
            col = rule['column']
            if col in res.columns:
                find_str = str(rule['find'])
                replace_str = str(rule['replace'])
                is_exact = rule.get('exact', False)
                
                if is_exact:
                    # Exact cell match
                    res[col] = res[col].replace(find_str, replace_str)
                else:
                    # Substring match (regex)
                    res[col] = res[col].astype(str).str.replace(find_str, replace_str, regex=True)
        return res

    @staticmethod
    def apply_expert_filters(df, options):
        """
        options: list of string keys mapped to expert functions.
        Supported options:
        - "trim_whitespace": 앞뒤 공백 제거
        - "remove_all_whitespace": 모든 공백 제거
        - "format_phone": 전화번호 하이픈(-) 포맷 통일 (010-XXXX-XXXX)
        - "drop_duplicates": (전체) 중복 행 제거
        - "drop_empty_rows": 모든 컬럼이 비어있는 행 제거
        - "to_upper": 영문 대문자로 통일
        - "to_lower": 영문 소문자로 통일
        - "mask_id": 주민/사업자등록번호 마스킹 (뒷부분 별표처리)
        - "extract_email": 문자열 내 이메일 형태만 추출
        - "remove_special_chars": 특수기호 제거 (문자, 숫자만 남김)
        - "normalize_numeric": 금액/숫자 콤마 제거 및 숫자화
        """
        res = df.copy()
        
        # Row level operations first
        if "drop_empty_rows" in options:
            res.dropna(how='all', inplace=True)
        if "drop_duplicates" in options:
            res.drop_duplicates(inplace=True)
            
        # Column iterations for text manipulations
        text_cols = res.select_dtypes(include=['object', 'string']).columns
        
        for col in text_cols:
            if "remove_all_whitespace" in options:
                res[col] = res[col].astype(str).str.replace(r'\s+', '', regex=True)
            elif "trim_whitespace" in options:
                res[col] = res[col].astype(str).str.strip()
                
            if "to_upper" in options:
                res[col] = res[col].astype(str).str.upper()
            if "to_lower" in options:
                res[col] = res[col].astype(str).str.lower()
                
            if "remove_special_chars" in options:
                res[col] = res[col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
                
            if "normalize_numeric" in options:
                # Remove commas, currency symbols, and common Korean suffixes
                res[col] = res[col].astype(str).str.replace(',', '').str.replace('₩', '').str.replace('원', '')
                # Try to convert to numeric where possible
                res[col] = pd.to_numeric(res[col], errors='ignore')
                
            if "format_phone" in options:
                def format_num(x):
                    s = re.sub(r'\D', '', str(x))
                    if len(s) == 11 and s.startswith('010'):
                        return f"{s[:3]}-{s[3:7]}-{s[7:]}"
                    elif len(s) == 10 and s.startswith('01'):
                        return f"{s[:3]}-{s[3:6]}-{s[6:]}"
                    return x
                res[col] = res[col].apply(format_num)
                
            if "mask_id" in options:
                # 주민등록번호: 6자리-7자리 -> 뒷 6자리 마스킹
                # 사업자: 3자리-2자리-5자리 -> 뒷 5자리 마스킹
                def mask_ids(val):
                    val = str(val)
                    val = re.sub(r'(\d{6})[-]?(\d)[0-9]{6}', r'\1-\2******', val)
                    val = re.sub(r'(\d{3})[-]?(\d{2})[-]?\d{5}', r'\1-\2-*****', val)
                    return val
                res[col] = res[col].apply(mask_ids)
                
            if "extract_email" in options:
                def get_email(x):
                    match = re.search(r'[\w\.-]+@[\w\.-]+', str(x))
                    return match.group(0) if match else x
                res[col] = res[col].apply(get_email)
                
        # Clean up possible "nan" string artifacts introduced by str conversions
        res.replace("nan", np.nan, inplace=True)
        res.replace("None", np.nan, inplace=True)
        
        return res

    @staticmethod
    def add_source_info(df, path):
        """Add a column with the source file path/name."""
        res = df.copy()
        res.insert(0, "원본경로", str(path))
        return res

    @staticmethod
    def perform_fuzzy_matching(df_left, df_right, key_col, threshold=85):
        """Perform approximate matching for columns that aren't exact matches."""
        left_unique = df_left[key_col].dropna().unique()
        right_unique = df_right[key_col].dropna().unique()
        
        # Build mapping
        mapping = {}
        for l_val in left_unique:
            # extractOne returns (match, score)
            best_match = process.extractOne(str(l_val), right_unique.astype(str), scorer=fuzz.token_sort_ratio)
            if best_match and best_match[1] >= threshold:
                mapping[l_val] = best_match[0]
        
        # Apply mapping to a temporary column
        df_temp = df_left.copy()
        df_temp['_fuzzy_key'] = df_temp[key_col].map(mapping)
        
        # Merge on fuzzy key
        res = pd.merge(df_temp, df_right, left_on='_fuzzy_key', right_on=key_col, how='left', suffixes=('', '_ref'))
        return res.drop(columns=['_fuzzy_key'])
