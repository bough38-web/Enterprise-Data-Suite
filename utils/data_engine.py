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
        return common[0] if common else df1.columns[0]

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
        """
        res = df.copy()
        
        # 1. Auto Target Filtering
        if filter_config.get('auto_target'):
            for col in ["시설구분", "요금구분"]:
                if col in res.columns:
                    res = res[res[col].astype(str).str.strip() == "대상"]
                    
        # 2. Custom Filters
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
                    
        return res

    @staticmethod
    def select_columns(df, columns, mode='keep'):
        if mode == 'keep':
            valid_cols = [c for c in columns if c in df.columns]
            return df[valid_cols] if valid_cols else df
        else:
            return df.drop(columns=[c for c in columns if c in df.columns], errors='ignore')

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
