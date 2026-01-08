"""
Base Rule Modules
Common rule implementations for business logic.
"""

import pandas as pd
from typing import Dict, Any, Optional


def calculate_growth_rate(data: pd.DataFrame, context: Optional[Dict] = None, 
                         current_col: str = "value", previous_col: str = "previous_value") -> float:
    """
    Calculate growth rate between current and previous values.
    
    Args:
        data: DataFrame containing the data
        context: Optional context dictionary
        current_col: Column name for current values
        previous_col: Column name for previous values
    
    Returns:
        Growth rate as percentage
    """
    current_sum = data[current_col].sum()
    previous_sum = data[previous_col].sum()
    
    if previous_sum == 0:
        return 0.0
    
    return ((current_sum - previous_sum) / previous_sum) * 100


def generate_performance_text(data: pd.DataFrame, context: Optional[Dict] = None,
                              value_col: str = "value", threshold: float = 0.0) -> str:
    """
    Generate performance text based on value and threshold.
    
    Args:
        data: DataFrame containing the data
        context: Optional context dictionary
        value_col: Column name for values
        threshold: Threshold for determining performance
    
    Returns:
        Performance text string
    """
    value = data[value_col].sum() if isinstance(data, pd.DataFrame) else data
    
    if value >= threshold:
        return "Performance Improved"
    else:
        return "Performance Declined"


def determine_color(value: float, threshold: float = 0.0,
                   positive_color: str = "#00FF00", negative_color: str = "#FF0000") -> str:
    """
    Determine color based on value and threshold.
    
    Args:
        value: Value to evaluate
        threshold: Threshold for color determination
        positive_color: Color for positive values
        negative_color: Color for negative values
    
    Returns:
        Color code string
    """
    return positive_color if value >= threshold else negative_color


def calculate_rankings(data: pd.DataFrame, context: Optional[Dict] = None,
                       value_col: str = "value", ascending: bool = False) -> pd.DataFrame:
    """
    Calculate rankings for values in a DataFrame.
    
    Args:
        data: DataFrame to rank
        context: Optional context dictionary
        value_col: Column name to rank by
        ascending: Whether to rank in ascending order
    
    Returns:
        DataFrame with rank column added
    """
    data = data.copy()
    data["rank"] = data[value_col].rank(ascending=ascending, method="dense").astype(int)
    return data


def filter_top_performers(data: pd.DataFrame, context: Optional[Dict] = None,
                          value_col: str = "value", n: int = 10) -> pd.DataFrame:
    """
    Filter top N performers.
    
    Args:
        data: DataFrame to filter
        context: Optional context dictionary
        value_col: Column name to sort by
        n: Number of top performers to return
    
    Returns:
        Filtered DataFrame
    """
    return data.nlargest(n, value_col)


def calculate_percentage_distribution(data: pd.DataFrame, context: Optional[Dict] = None,
                                     value_col: str = "value", group_col: str = "category") -> pd.DataFrame:
    """
    Calculate percentage distribution across groups.
    
    Args:
        data: DataFrame to process
        context: Optional context dictionary
        value_col: Column name for values
        group_col: Column name for grouping
    
    Returns:
        DataFrame with percentage distribution
    """
    grouped = data.groupby(group_col)[value_col].sum().reset_index()
    total = grouped[value_col].sum()
    grouped["percentage"] = (grouped[value_col] / total * 100).round(2)
    return grouped


def detect_trend(data: pd.DataFrame, context: Optional[Dict] = None,
                value_col: str = "value", date_col: str = "date") -> str:
    """
    Detect trend direction (increasing, decreasing, stable).
    
    Args:
        data: DataFrame with time series data
        context: Optional context dictionary
        value_col: Column name for values
        date_col: Column name for dates
    
    Returns:
        Trend direction string
    """
    if len(data) < 2:
        return "insufficient_data"
    
    data = data.sort_values(date_col)
    values = data[value_col].values
    
    # Calculate trend
    first_half = values[:len(values)//2].mean()
    second_half = values[len(values)//2:].mean()
    
    change_pct = ((second_half - first_half) / first_half * 100) if first_half != 0 else 0
    
    if abs(change_pct) < 5:
        return "stable"
    elif change_pct > 0:
        return "increasing"
    else:
        return "decreasing"


def format_number(value: float, format_type: str = "comma", decimals: int = 2) -> str:
    """
    Format a number according to specified format.
    
    Args:
        value: Number to format
        format_type: Format type ('comma', 'currency', 'percentage')
        decimals: Number of decimal places
    
    Returns:
        Formatted string
    """
    if format_type == "comma":
        return f"{value:,.{decimals}f}"
    elif format_type == "currency":
        return f"${value:,.{decimals}f}"
    elif format_type == "percentage":
        return f"{value:.{decimals}f}%"
    else:
        return str(value)

