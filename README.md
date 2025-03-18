# Survey Aggregator and HTML Generator

This project aggregates survey data from multiple Excel files and generates an HTML file displaying:

- Aggregated Regular Votes
- Aggregated Integer Votes
- Aggregated Comments (with text wrapping)
- A generated plot image (embedded using base64)

## Requirements

- **Python 3.x**

### Python Libraries

The following Python libraries are required:

- **pandas** (for data manipulation)
- **matplotlib** (for plotting)
- **openpyxl** (for reading Excel `.xlsx` files)

You can install the necessary libraries using pip:

```bash
pip install pandas matplotlib openpyxl
