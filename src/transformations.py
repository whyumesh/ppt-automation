"""
Data Transformations
Reusable transformation functions for aggregations, calculations, and filters.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Callable, Tuple
from datetime import datetime, timedelta


class DataTransformations:
    """Collection of reusable data transformation functions."""
    
    @staticmethod
    def aggregate(df: pd.DataFrame, 
                  group_by: List[str],
                  agg_functions: Dict[str, List[str]]) -> pd.DataFrame:
        """
        Aggregate data by grouping columns.
        
        Args:
            df: DataFrame to aggregate
            group_by: List of columns to group by
            agg_functions: Dictionary mapping columns to aggregation functions
        
        Returns:
            Aggregated DataFrame
        """
        return df.groupby(group_by).agg(agg_functions).reset_index()
    
    @staticmethod
    def calculate_percentage(df: pd.DataFrame,
                            numerator_col: str,
                            denominator_col: str,
                            output_col: str = "percentage") -> pd.DataFrame:
        """
        Calculate percentage from two columns.
        
        Args:
            df: DataFrame to process
            numerator_col: Column name for numerator
            denominator_col: Column name for denominator
            output_col: Name of output column
        
        Returns:
            DataFrame with percentage column added
        """
        df = df.copy()
        df[output_col] = (df[numerator_col] / df[denominator_col] * 100).fillna(0)
        return df
    
    @staticmethod
    def calculate_delta(df: pd.DataFrame,
                        value_col: str,
                        compare_col: str,
                        output_col: str = "delta") -> pd.DataFrame:
        """
        Calculate delta (difference) between two columns.
        
        Args:
            df: DataFrame to process
            value_col: Column name for current value
            compare_col: Column name for comparison value
            output_col: Name of output column
        
        Returns:
            DataFrame with delta column added
        """
        df = df.copy()
        df[output_col] = df[value_col] - df[compare_col]
        return df
    
    @staticmethod
    def calculate_percentage_change(df: pd.DataFrame,
                                   current_col: str,
                                   previous_col: str,
                                   output_col: str = "pct_change") -> pd.DataFrame:
        """
        Calculate percentage change between two columns.
        
        Args:
            df: DataFrame to process
            current_col: Column name for current value
            previous_col: Column name for previous value
            output_col: Name of output column
        
        Returns:
            DataFrame with percentage change column added
        """
        df = df.copy()
        df[output_col] = ((df[current_col] - df[previous_col]) / df[previous_col] * 100).fillna(0)
        return df
    
    @staticmethod
    def calculate_rank(df: pd.DataFrame,
                      value_col: str,
                      output_col: str = "rank",
                      ascending: bool = False,
                      method: str = "dense") -> pd.DataFrame:
        """
        Calculate rankings for a column.
        
        Args:
            df: DataFrame to process
            value_col: Column name to rank
            output_col: Name of output column
            ascending: Whether to rank in ascending order
            method: Ranking method ('dense', 'min', 'max', 'first')
        
        Returns:
            DataFrame with rank column added
        """
        df = df.copy()
        df[output_col] = df[value_col].rank(ascending=ascending, method=method).astype(int)
        return df
    
    @staticmethod
    def filter_by_threshold(df: pd.DataFrame,
                           column: str,
                           threshold: float,
                           operator: str = ">=") -> pd.DataFrame:
        """
        Filter rows based on threshold.
        
        Args:
            df: DataFrame to filter
            column: Column name to filter on
            threshold: Threshold value
            operator: Comparison operator ('>', '>=', '<', '<=', '==', '!=')
        
        Returns:
            Filtered DataFrame
        """
        operators = {
            ">": lambda x, y: x > y,
            ">=": lambda x, y: x >= y,
            "<": lambda x, y: x < y,
            "<=": lambda x, y: x <= y,
            "==": lambda x, y: x == y,
            "!=": lambda x, y: x != y
        }
        
        if operator not in operators:
            raise ValueError(f"Invalid operator: {operator}")
        
        mask = operators[operator](df[column], threshold)
        return df[mask].copy()
    
    @staticmethod
    def filter_top_n(df: pd.DataFrame,
                    value_col: str,
                    n: int,
                    ascending: bool = False) -> pd.DataFrame:
        """
        Filter top N rows by value.
        
        Args:
            df: DataFrame to filter
            value_col: Column name to sort by
            n: Number of rows to return
            ascending: Whether to sort in ascending order
        
        Returns:
            Filtered DataFrame
        """
        return df.nlargest(n, value_col) if not ascending else df.nsmallest(n, value_col)
    
    @staticmethod
    def filter_by_condition(df: pd.DataFrame,
                           condition: Callable[[pd.DataFrame], pd.Series]) -> pd.DataFrame:
        """
        Filter rows using a custom condition function.
        
        Args:
            df: DataFrame to filter
            condition: Function that takes DataFrame and returns boolean Series
        
        Returns:
            Filtered DataFrame
        """
        mask = condition(df)
        return df[mask].copy()
    
    @staticmethod
    def round_values(df: pd.DataFrame,
                    columns: List[str],
                    decimals: int = 2) -> pd.DataFrame:
        """
        Round numeric values in specified columns.
        
        Args:
            df: DataFrame to process
            columns: List of column names to round
            decimals: Number of decimal places
        
        Returns:
            DataFrame with rounded values
        """
        df = df.copy()
        for col in columns:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].round(decimals)
        return df
    
    @staticmethod
    def format_numbers(df: pd.DataFrame,
                      columns: List[str],
                      format_type: str = "comma") -> pd.DataFrame:
        """
        Format numbers in specified columns (for display purposes).
        
        Args:
            df: DataFrame to process
            columns: List of column names to format
            format_type: Format type ('comma', 'currency', 'percentage')
        
        Returns:
            DataFrame with formatted values (as strings)
        """
        df = df.copy()
        for col in columns:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                if format_type == "comma":
                    df[col] = df[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
                elif format_type == "currency":
                    df[col] = df[col].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
                elif format_type == "percentage":
                    df[col] = df[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "")
        return df
    
    @staticmethod
    def pivot_table(df: pd.DataFrame,
                   index: List[str],
                   columns: List[str],
                   values: List[str],
                   aggfunc: str = "sum") -> pd.DataFrame:
        """
        Create a pivot table.
        
        Args:
            df: DataFrame to pivot
            index: List of columns for index
            columns: List of columns for columns
            values: List of columns for values
            aggfunc: Aggregation function
        
        Returns:
            Pivoted DataFrame
        """
        return pd.pivot_table(df, index=index, columns=columns, values=values, aggfunc=aggfunc)
    
    @staticmethod
    def merge_dataframes(df1: pd.DataFrame,
                        df2: pd.DataFrame,
                        on: List[str],
                        how: str = "inner") -> pd.DataFrame:
        """
        Merge two DataFrames.
        
        Args:
            df1: First DataFrame
            df2: Second DataFrame
            on: List of columns to merge on
            how: Merge type ('inner', 'outer', 'left', 'right')
        
        Returns:
            Merged DataFrame
        """
        return pd.merge(df1, df2, on=on, how=how)
    
    @staticmethod
    def calculate_trend(df: pd.DataFrame,
                       date_col: str,
                       value_col: str,
                       period: str = "month") -> pd.DataFrame:
        """
        Calculate trends over time periods.
        
        Args:
            df: DataFrame to process
            date_col: Column name containing dates
            value_col: Column name containing values
            period: Time period ('day', 'week', 'month', 'quarter', 'year')
        
        Returns:
            DataFrame with trend calculations
        """
        df = df.copy()
        df[date_col] = pd.to_datetime(df[date_col])
        
        # Group by period
        if period == "month":
            df["period"] = df[date_col].dt.to_period("M")
        elif period == "quarter":
            df["period"] = df[date_col].dt.to_period("Q")
        elif period == "year":
            df["period"] = df[date_col].dt.to_period("Y")
        else:
            df["period"] = df[date_col]
        
        # Calculate period aggregates
        trend_df = df.groupby("period")[value_col].agg(['sum', 'mean', 'count']).reset_index()
        trend_df.columns = ["period", f"{value_col}_sum", f"{value_col}_mean", f"{value_col}_count"]
        
        # Calculate period-over-period change
        trend_df[f"{value_col}_pct_change"] = trend_df[f"{value_col}_sum"].pct_change() * 100
        
        return trend_df
    
    @staticmethod
    def apply_transformation_pipeline(df: pd.DataFrame,
                                     transformations: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Apply a series of transformations in sequence.
        
        Args:
            df: DataFrame to process
            transformations: List of transformation dictionaries
        
        Returns:
            Transformed DataFrame
        """
        result_df = df.copy()
        
        for transform in transformations:
            transform_type = transform.get("type")
            params = transform.get("params", {})
            
            if transform_type == "aggregate":
                result_df = DataTransformations.aggregate(result_df, **params)
            elif transform_type == "filter_threshold":
                result_df = DataTransformations.filter_by_threshold(result_df, **params)
            elif transform_type == "filter_top_n":
                result_df = DataTransformations.filter_top_n(result_df, **params)
            elif transform_type == "calculate_percentage":
                result_df = DataTransformations.calculate_percentage(result_df, **params)
            elif transform_type == "calculate_delta":
                result_df = DataTransformations.calculate_delta(result_df, **params)
            elif transform_type == "round":
                result_df = DataTransformations.round_values(result_df, **params)
            elif transform_type == "format":
                result_df = DataTransformations.format_numbers(result_df, **params)
            else:
                print(f"Warning: Unknown transformation type: {transform_type}")
        
        return result_df


if __name__ == "__main__":
    # Example usage
    import pandas as pd
    
    # Create sample data
    data = {
        'category': ['A', 'A', 'B', 'B', 'C', 'C'],
        'value': [10, 20, 30, 40, 50, 60],
        'previous_value': [8, 18, 25, 35, 45, 55]
    }
    df = pd.DataFrame(data)
    
    print("Original DataFrame:")
    print(df)
    print()
    
    # Calculate percentage change
    df_pct = DataTransformations.calculate_percentage_change(df, 'value', 'previous_value')
    print("With percentage change:")
    print(df_pct)
    print()
    
    # Filter top 3
    df_top = DataTransformations.filter_top_n(df, 'value', 3)
    print("Top 3 by value:")
    print(df_top)

