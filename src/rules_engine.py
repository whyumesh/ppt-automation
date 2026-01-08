"""
Business Rules Engine
Core rule evaluation engine for implementing business logic without AI dependencies.
"""

import yaml
import pandas as pd
from typing import Dict, List, Any, Optional, Callable
from pathlib import Path
import importlib
import os


class RulesEngine:
    """Evaluates business rules based on configuration and data."""
    
    def __init__(self, rules_config: Optional[str] = None):
        """
        Initialize the rules engine.
        
        Args:
            rules_config: Optional path to rules configuration YAML file
        """
        self.rules_config = rules_config
        self.rules = {}
        self.rule_modules = {}
        
        if rules_config and os.path.exists(rules_config):
            self._load_rules()
        else:
            # Initialize empty rules if config doesn't exist or is empty
            self.rules = {}
    
    def _load_rules(self):
        """Load rules from configuration file."""
        try:
            with open(self.rules_config, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
                if config is None:
                    config = {}
                self.rules = config.get("rules", {})
                if self.rules is None:
                    self.rules = {}
        except Exception as e:
            print(f"Warning: Could not load rules from {self.rules_config}: {e}")
            self.rules = {}
    
    def evaluate_rule(self, rule_name: str, data: Any, context: Optional[Dict] = None) -> Any:
        """
        Evaluate a specific rule.
        
        Args:
            rule_name: Name of the rule to evaluate
            data: Data to evaluate the rule against
            context: Optional context dictionary
        
        Returns:
            Result of rule evaluation
        """
        if rule_name not in self.rules:
            raise ValueError(f"Rule '{rule_name}' not found in configuration")
        
        rule_def = self.rules[rule_name]
        rule_type = rule_def.get("type")
        
        if rule_type == "calculation":
            return self._evaluate_calculation_rule(rule_def, data, context)
        elif rule_type == "filter":
            return self._evaluate_filter_rule(rule_def, data, context)
        elif rule_type == "formatting":
            return self._evaluate_formatting_rule(rule_def, data, context)
        elif rule_type == "conditional":
            return self._evaluate_conditional_rule(rule_def, data, context)
        elif rule_type == "text_generation":
            return self._evaluate_text_generation_rule(rule_def, data, context)
        elif rule_type == "custom":
            return self._evaluate_custom_rule(rule_def, data, context)
        else:
            raise ValueError(f"Unknown rule type: {rule_type}")
    
    def _evaluate_calculation_rule(self, rule_def: Dict, data: Any, context: Optional[Dict]) -> Any:
        """Evaluate a calculation rule."""
        operation = rule_def.get("operation")
        params = rule_def.get("params", {})
        
        if operation == "sum":
            column = params.get("column")
            return data[column].sum() if isinstance(data, pd.DataFrame) else sum(data)
        
        elif operation == "mean":
            column = params.get("column")
            return data[column].mean() if isinstance(data, pd.DataFrame) else sum(data) / len(data)
        
        elif operation == "count":
            column = params.get("column")
            if isinstance(data, pd.DataFrame):
                return len(data) if column is None else data[column].count()
            return len(data)
        
        elif operation == "percentage":
            numerator = params.get("numerator")
            denominator = params.get("denominator")
            if isinstance(data, pd.DataFrame):
                return (data[numerator].sum() / data[denominator].sum() * 100) if data[denominator].sum() != 0 else 0
            return (numerator / denominator * 100) if denominator != 0 else 0
        
        elif operation == "delta":
            current = params.get("current")
            previous = params.get("previous")
            if isinstance(data, pd.DataFrame):
                return data[current].sum() - data[previous].sum()
            return current - previous
        
        elif operation == "percentage_change":
            current = params.get("current")
            previous = params.get("previous")
            if isinstance(data, pd.DataFrame):
                current_sum = data[current].sum()
                previous_sum = data[previous].sum()
                return ((current_sum - previous_sum) / previous_sum * 100) if previous_sum != 0 else 0
            return ((current - previous) / previous * 100) if previous != 0 else 0
        
        else:
            raise ValueError(f"Unknown calculation operation: {operation}")
    
    def _evaluate_filter_rule(self, rule_def: Dict, data: pd.DataFrame, context: Optional[Dict]) -> pd.DataFrame:
        """Evaluate a filter rule."""
        if not isinstance(data, pd.DataFrame):
            raise ValueError("Filter rules require a DataFrame")
        
        filter_type = rule_def.get("filter_type")
        params = rule_def.get("params", {})
        
        if filter_type == "threshold":
            column = params.get("column")
            threshold = params.get("threshold")
            operator = params.get("operator", ">=")
            
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
            
            mask = operators[operator](data[column], threshold)
            return data[mask].copy()
        
        elif filter_type == "top_n":
            column = params.get("column")
            n = params.get("n", 10)
            ascending = params.get("ascending", False)
            return data.nlargest(n, column) if not ascending else data.nsmallest(n, column)
        
        elif filter_type == "contains":
            column = params.get("column")
            value = params.get("value")
            return data[data[column].str.contains(value, na=False)].copy()
        
        elif filter_type == "in_list":
            column = params.get("column")
            values = params.get("values", [])
            return data[data[column].isin(values)].copy()
        
        else:
            raise ValueError(f"Unknown filter type: {filter_type}")
    
    def _evaluate_formatting_rule(self, rule_def: Dict, data: Any, context: Optional[Dict]) -> Dict[str, Any]:
        """Evaluate a formatting rule."""
        format_type = rule_def.get("format_type")
        params = rule_def.get("params", {})
        
        result = {}
        
        if format_type == "round":
            value = params.get("value")
            decimals = params.get("decimals", 2)
            result["formatted_value"] = round(value, decimals)
        
        elif format_type == "percentage":
            value = params.get("value")
            decimals = params.get("decimals", 1)
            result["formatted_value"] = f"{value:.{decimals}f}%"
        
        elif format_type == "currency":
            value = params.get("value")
            result["formatted_value"] = f"${value:,.2f}"
        
        elif format_type == "color":
            value = params.get("value")
            threshold = params.get("threshold", 0)
            positive_color = params.get("positive_color", "#00FF00")
            negative_color = params.get("negative_color", "#FF0000")
            
            result["color"] = positive_color if value >= threshold else negative_color
        
        return result
    
    def _evaluate_conditional_rule(self, rule_def: Dict, data: Any, context: Optional[Dict]) -> Dict[str, Any]:
        """Evaluate a conditional rule."""
        condition = rule_def.get("condition")
        true_action = rule_def.get("true_action", {})
        false_action = rule_def.get("false_action", {})
        
        # Evaluate condition
        condition_result = self._evaluate_condition(condition, data, context)
        
        # Return appropriate action result
        if condition_result:
            return self._execute_action(true_action, data, context)
        else:
            return self._execute_action(false_action, data, context)
    
    def _evaluate_condition(self, condition: Dict, data: Any, context: Optional[Dict]) -> bool:
        """Evaluate a condition expression."""
        condition_type = condition.get("type")
        
        if condition_type == "compare":
            left = self._get_value(condition.get("left"), data, context)
            right = self._get_value(condition.get("right"), data, context)
            operator = condition.get("operator", "==")
            
            operators = {
                ">": lambda x, y: x > y,
                ">=": lambda x, y: x >= y,
                "<": lambda x, y: x < y,
                "<=": lambda x, y: x <= y,
                "==": lambda x, y: x == y,
                "!=": lambda x, y: x != y
            }
            
            return operators[operator](left, right)
        
        elif condition_type == "and":
            conditions = condition.get("conditions", [])
            return all(self._evaluate_condition(c, data, context) for c in conditions)
        
        elif condition_type == "or":
            conditions = condition.get("conditions", [])
            return any(self._evaluate_condition(c, data, context) for c in conditions)
        
        elif condition_type == "not":
            sub_condition = condition.get("condition")
            return not self._evaluate_condition(sub_condition, data, context)
        
        else:
            raise ValueError(f"Unknown condition type: {condition_type}")
    
    def _get_value(self, value_spec: Any, data: Any, context: Optional[Dict]) -> Any:
        """Get a value from data or context."""
        if isinstance(value_spec, dict):
            value_type = value_spec.get("type")
            if value_type == "data_column":
                column = value_spec.get("column")
                if isinstance(data, pd.DataFrame):
                    return data[column].sum() if value_spec.get("aggregate") == "sum" else data[column].iloc[0]
            elif value_type == "context":
                key = value_spec.get("key")
                return context.get(key) if context else None
            elif value_type == "literal":
                return value_spec.get("value")
        else:
            return value_spec
    
    def _execute_action(self, action: Dict, data: Any, context: Optional[Dict]) -> Dict[str, Any]:
        """Execute an action."""
        action_type = action.get("type")
        
        if action_type == "set_value":
            return {"value": action.get("value")}
        
        elif action_type == "set_text":
            template = action.get("template", "")
            # Simple template substitution
            if context:
                for key, value in context.items():
                    template = template.replace(f"{{{key}}}", str(value))
            return {"text": template}
        
        elif action_type == "set_color":
            return {"color": action.get("color")}
        
        elif action_type == "evaluate_rule":
            rule_name = action.get("rule")
            return self.evaluate_rule(rule_name, data, context)
        
        else:
            return {}
    
    def _evaluate_text_generation_rule(self, rule_def: Dict, data: Any, context: Optional[Dict]) -> str:
        """Evaluate a text generation rule."""
        template = rule_def.get("template", "")
        params = rule_def.get("params", {})
        
        # Get values from data or context
        values = {}
        for key, value_spec in params.items():
            values[key] = self._get_value(value_spec, data, context)
        
        # Substitute values in template
        result = template
        for key, value in values.items():
            result = result.replace(f"{{{key}}}", str(value))
        
        return result
    
    def _evaluate_custom_rule(self, rule_def: Dict, data: Any, context: Optional[Dict]) -> Any:
        """Evaluate a custom rule from a module."""
        module_name = rule_def.get("module")
        function_name = rule_def.get("function")
        params = rule_def.get("params", {})
        
        # Load module if not already loaded
        if module_name not in self.rule_modules:
            try:
                module = importlib.import_module(f"src.rules.{module_name}")
                self.rule_modules[module_name] = module
            except ImportError:
                raise ImportError(f"Could not import rule module: {module_name}")
        
        module = self.rule_modules[module_name]
        
        if not hasattr(module, function_name):
            raise AttributeError(f"Function '{function_name}' not found in module '{module_name}'")
        
        function = getattr(module, function_name)
        return function(data, context, **params)
    
    def evaluate_all_rules(self, data: Dict[str, pd.DataFrame], context: Optional[Dict] = None) -> Dict[str, Any]:
        """
        Evaluate all rules for given data.
        
        Args:
            data: Dictionary mapping data source names to DataFrames
            context: Optional context dictionary
        
        Returns:
            Dictionary mapping rule names to evaluation results
        """
        results = {}
        
        # Ensure rules is a dict
        if self.rules is None:
            self.rules = {}
        
        # If no rules configured, return empty results
        if not self.rules:
            return results
        
        for rule_name in self.rules.keys():
            try:
                # Determine which data source to use for this rule
                rule_def = self.rules[rule_name]
                data_source = rule_def.get("data_source")
                
                rule_data = data.get(data_source) if data_source else list(data.values())[0] if data else None
                
                if rule_data is not None:
                    results[rule_name] = self.evaluate_rule(rule_name, rule_data, context)
                else:
                    results[rule_name] = None
            except Exception as e:
                results[rule_name] = {"error": str(e)}
        
        return results


if __name__ == "__main__":
    # Example usage
    import sys
    import pandas as pd
    
    if len(sys.argv) < 2:
        print("Usage: python rules_engine.py <rules_config>")
        sys.exit(1)
    
    rules_config = sys.argv[1]
    
    # Create sample data
    data = {
        'value': [10, 20, 30, 40, 50],
        'previous_value': [8, 18, 25, 35, 45]
    }
    df = pd.DataFrame(data)
    
    engine = RulesEngine(rules_config=rules_config)
    results = engine.evaluate_all_rules({"main": df})
    
    print("Rule evaluation results:")
    for rule_name, result in results.items():
        print(f"  {rule_name}: {result}")

