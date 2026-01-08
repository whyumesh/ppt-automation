# Automated PowerPoint Deck Creation from Excel

A rule-based, fully local automation system for generating PowerPoint decks from Excel files. No AI/ML models required - 100% deterministic and transparent.

## Features

- **Rule-Based**: Pure Python logic with no AI dependencies
- **Fully Local**: All processing happens on your machine
- **Data Privacy Compliant**: No cloud services or external APIs
- **Configurable**: YAML-based configuration for easy maintenance
- **Extensible**: Modular architecture for adding new rules and slides

## Project Structure

```
ppt-automation/
├── src/                    # Source code
│   ├── template_extractor.py    # Extract PPT templates
│   ├── excel_analyzer.py         # Analyze Excel files
│   ├── rule_discoverer.py        # Discover business rules
│   ├── data_loader.py            # Load Excel files
│   ├── data_normalizer.py        # Normalize data
│   ├── transformations.py        # Data transformations
│   ├── rules_engine.py           # Business rules engine
│   ├── rules/                    # Rule modules
│   ├── ppt_generator.py          # PPT generation orchestrator
│   ├── ppt_builder.py            # Slide building utilities
│   ├── ppt_formatter.py          # Formatting utilities
│   └── validator.py              # Validation utilities
├── config/                 # Configuration files
│   ├── slides.yaml        # Slide mappings
│   ├── rules.yaml         # Business rules
│   ├── formatting.yaml    # Formatting rules
│   └── schema.yaml        # Excel schemas
├── templates/             # PPT templates
├── docs/                  # Documentation
├── tests/                 # Unit tests
├── validation/            # Validation results
├── Data/                  # Historical data (unchanged)
├── requirements.txt       # Python dependencies
└── main.py               # Main entry point
```

## Installation

1. Install Python 3.8 or higher

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### 1. Reverse Engineering (Discover Rules)

First, analyze historical Excel and PPT files to discover mappings and rules:

```bash
python main.py analyze "Data/Apr 2025/AIL LT Working file.xlsx" "Data/Apr 2025/AIL LT - April'25.pptx" --output-dir analysis
```

This will generate:
- `analysis/template_info.json` - PPT template structure
- `analysis/excel_info.json` - Excel file structure
- `analysis/discovered_rules.json` - Discovered business rules

### 2. Configure Mappings

Based on the analysis results, update the configuration files:

- **config/slides.yaml**: Map Excel data sources to PPT slides
- **config/rules.yaml**: Define business rules and calculations
- **config/formatting.yaml**: Define formatting rules
- **config/schema.yaml**: Define Excel schemas

### 3. Generate PowerPoint Deck

Generate a PPT deck from Excel files:

```bash
python main.py generate "Data/Apr 2025" "output/Apr_2025_Generated.pptx" --template "templates/template.pptx"
```

### 4. Validate Output

Compare generated PPT with manual version:

```bash
python -m src.validator "Data/Apr 2025/AIL LT - April'25.pptx" "output/Apr_2025_Generated.pptx" "validation/report.json"
```

## Configuration

### Slide Mappings (config/slides.yaml)

Define how Excel data maps to PowerPoint slides:

```yaml
slides:
  - slide_number: 1
    slide_type: "title"
    title: "Monthly Report"
    subtitle: "April 2025"
  
  - slide_number: 2
    slide_type: "table"
    title: "Summary Data"
    table_mapping:
      data_source: "working_file"
      sheet: "Summary"
      columns: ["Category", "Value"]
      filters:
        - column: "Value"
          operator: ">="
          value: 0
```

### Business Rules (config/rules.yaml)

Define calculation and transformation rules:

```yaml
rules:
  calculate_growth:
    type: "calculation"
    operation: "percentage_change"
    params:
      current: "current_value"
      previous: "previous_value"
    data_source: "main_data"
```

### Formatting Rules (config/formatting.yaml)

Define formatting styles:

```yaml
formatting:
  fonts:
    default_size: 12
    title_size: 24
  colors:
    positive: "#00FF00"
    negative: "#FF0000"
```

## Workflow

1. **Reverse Engineering Phase**: Analyze historical data to discover rules
2. **Configuration Phase**: Document discovered rules in YAML configs
3. **Data Processing**: Load and normalize Excel files
4. **Rule Application**: Apply business rules to processed data
5. **PPT Generation**: Generate PowerPoint deck from processed data
6. **Validation**: Compare with manual versions and refine rules

## Key Components

### Data Processing Layer
- **DataLoader**: Loads Excel files (.xlsx, .xlsb)
- **DataNormalizer**: Normalizes column names and data types
- **Transformations**: Applies aggregations, calculations, filters

### Business Rules Engine
- **RulesEngine**: Evaluates rules based on configuration
- **Rule Modules**: Custom rule implementations in `src/rules/`

### PPT Generation Layer
- **PPTGenerator**: Main orchestrator
- **PPTBuilder**: Builds slides and populates content
- **PPTFormatter**: Applies formatting (fonts, colors, alignment)

### Validation
- **PPTValidator**: Compares generated PPTs with manual versions
- Generates detailed validation reports

## Example: Processing a Month

```python
from main import PPTPipeline

pipeline = PPTPipeline(
    config_dir="config",
    template_path="templates/template.pptx"
)

pipeline.process_month(
    month_data_dir="Data/Apr 2025",
    output_path="output/Apr_2025_Generated.pptx"
)
```

## Extending the System

### Adding New Rules

1. Create a new rule module in `src/rules/`:
```python
def my_custom_rule(data, context, **params):
    # Your rule logic here
    return result
```

2. Reference it in `config/rules.yaml`:
```yaml
my_rule:
  type: "custom"
  module: "my_rules"
  function: "my_custom_rule"
  params:
    param1: value1
```

### Adding New Slide Types

1. Add slide generation method to `PPTGenerator`:
```python
def _generate_custom_slide(self, slide, slide_config, data):
    # Your slide generation logic
    pass
```

2. Add mapping in `config/slides.yaml`:
```yaml
- slide_number: X
  slide_type: "custom"
  # ... configuration
```

## Requirements

- Python 3.8+
- pandas >= 2.0.0
- numpy >= 1.24.0
- python-pptx >= 0.6.21
- openpyxl >= 3.1.0
- pyxlsb >= 1.0.10
- pyyaml >= 6.0
- python-dateutil >= 2.8.2

## License

[Your License Here]

## Support

For issues or questions, please refer to the documentation in `docs/` or create an issue.

