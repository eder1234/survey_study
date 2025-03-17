# Aggregate Survey Votes

## Description
This script aggregates survey votes from multiple Excel (`.xlsx`) files in a specified folder, computes the cumulative number of votes per question and per class, and generates both a CSV file and a graphical representation (`results.png`) displaying the percentage distribution of votes.

## Requirements
Ensure you have Python installed along with the following libraries:

```bash
pip install pandas matplotlib openpyxl
```

## Folder Structure
Your working directory should contain:

```
.
├── main.py
└── surveys/
    ├── file1.xlsx
    ├── file2.xlsx
    └── ...
```

Replace `surveys` with your actual directory name.

## How to Use

1. **Place all `.xlsx` files** (containing the survey data) into your chosen directory (e.g., `surveys`).

2. **Run the script:**

```bash
python main.py
```

- The script searches for the question starting with the provided text:

```python
question_text = "Le formateur a communiqué de façon dynamique"
```

(You may change this text if needed to match exactly your data.)

3. **Output Files**:
   - `aggregated_votes.csv`: Contains cumulative votes.
   - `results.png`: Displays a percentage-based visual representation of votes per question and class.

## File Structure
```
.
├── your_script.py
├── surveys
│   ├── file1.xlsx
│   ├── file2.xlsx
│   └── ...
├── aggregated_votes.csv
└── results.png
```

## Troubleshooting
- If you encounter a Qt-related issue (particularly on headless Linux environments), the script explicitly sets the Matplotlib backend to `'Agg'`, suitable for environments without graphical interfaces.

## License
Feel free to modify and distribute this script as needed.


