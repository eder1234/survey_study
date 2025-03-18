import pandas as pd
import glob
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import base64

# Function to find question start position
def find_question_start(df, question_text):
    location = df.isin([question_text])
    result = location.stack().idxmax()
    row, col_name = result
    col = df.columns.get_loc(col_name)
    return row, col

# Extraction function for regular votes
def extract_regular_votes(df, start_row, start_col, num_questions=17, classes_per_question=5, row_step=3):
    votes_data, questions = [], []
    for q in range(num_questions):
        question_row = start_row + q * row_step
        votes_row = question_row + 1
        question_text = df.iloc[question_row, start_col].replace(',', ' ')
        votes = df.iloc[votes_row, start_col+1:start_col+1+classes_per_question].values.astype(float)
        questions.append(question_text)
        votes_data.append(votes)
    class_labels = ['Tout à fait d\'accord', 'D\'accord', 'Indécis', 'En désaccord', 'Tout à fait en désaccord']
    return pd.DataFrame(votes_data, columns=class_labels, index=questions)

# Extraction function for the last question with integer classes (0-10)
def extract_integer_votes(df, last_question_row, start_col, num_classes=11):
    question_text = df.iloc[last_question_row, start_col].replace(',', ' ')
    votes = df.iloc[last_question_row + 1, start_col+1:start_col+1+num_classes].values.astype(float)
    class_labels = [str(i) for i in range(num_classes)]
    return pd.DataFrame([votes], columns=class_labels, index=[question_text])

# Extraction function for comments
def extract_comments(df):
    comments_location = df.isin(["Commentaires :"])
    if not comments_location.any().any():
        return []
    comments_row, comments_col_name = comments_location.stack().idxmax()
    comments_col = df.columns.get_loc(comments_col_name)
    comments = []
    for row in range(comments_row, len(df)):
        cell_value = df.iloc[row, comments_col + 1]
        if pd.notnull(cell_value):
            comment = str(cell_value).replace("<br />", " ").strip()
            if comment:
                comments.append(comment)
    return comments

# Main aggregation function and HTML generation
def aggregate_votes_and_generate_html(folder_path, question_text):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    cumulative_regular_votes, cumulative_integer_votes, all_comments = None, None, []

    for file in all_files:
        df = pd.read_excel(file)
        start_row, start_col = find_question_start(df, question_text)
        regular_votes = extract_regular_votes(df, start_row, start_col)
        integer_votes = extract_integer_votes(df, start_row + 17*3, start_col)
        comments = extract_comments(df)
        cumulative_regular_votes = regular_votes if cumulative_regular_votes is None else cumulative_regular_votes + regular_votes
        cumulative_integer_votes = integer_votes if cumulative_integer_votes is None else cumulative_integer_votes + integer_votes
        for comment in comments:
            # Exclude file info for HTML display; just keep the comment text.
            all_comments.append({'Comment': comment})

    # Save aggregated CSV files (optional)
    regular_votes_path = os.path.join(folder_path, 'aggregated_regular_votes.csv')
    integer_votes_path = os.path.join(folder_path, 'aggregated_integer_votes.csv')
    comments_path = os.path.join(folder_path, 'aggregated_comments.csv')
    cumulative_regular_votes.to_csv(regular_votes_path, encoding='utf-8-sig')
    cumulative_integer_votes.to_csv(integer_votes_path, encoding='utf-8-sig')
    pd.DataFrame(all_comments).to_csv(comments_path, encoding='utf-8-sig', index=False)
    print(f"Aggregated regular votes saved to {regular_votes_path}")
    print(f"Aggregated integer votes saved to {integer_votes_path}")
    print(f"Aggregated comments saved to {comments_path}")

    # Generate the votes plot image
    df_percentages = cumulative_regular_votes.div(cumulative_regular_votes.sum(axis=1), axis=0) * 100
    fig, ax = plt.subplots(figsize=(12, 8))
    df_percentages.iloc[::-1].plot(kind='barh', stacked=True, ax=ax, colormap='RdYlGn')
    plt.xlabel('Percentage (%)')
    plt.ylabel('Question')
    plt.title('Percentage of Votes per Class (Regular Questions)')
    plt.legend(title='Vote Type', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig("results.png")
    plt.close()
    print("Plot saved to results.png")

    # Encode the image as base64 to embed directly into HTML
    with open("results.png", "rb") as img_file:
        encoded_img = base64.b64encode(img_file.read()).decode('utf-8')

    # Compute integer vote percentages (assuming one row)
    integer_votes_percentages = cumulative_integer_votes.div(cumulative_integer_votes.sum(axis=1), axis=0) * 100

    # Create HTML file to display the aggregated results and image
    html_file_path = os.path.join(folder_path, 'aggregated_results.html')
    with open(html_file_path, 'w', encoding='utf-8') as f:
        f.write("<html><head><meta charset='utf-8'>")
        f.write("<title>Aggregated Survey Results</title>")
        # CSS for better table formatting and word-wrapping
        f.write("""
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            table { border-collapse: collapse; width: 100%; table-layout: fixed; }
            th, td { border: 1px solid #ddd; padding: 8px; word-wrap: break-word; }
            th { background-color: #f2f2f2; }
            img { max-width: 100%; height: auto; }
        </style>
        """)
        f.write("</head><body>")
        f.write("<h1>Aggregated Survey Results</h1>")

        # Regular votes table
        f.write("<h2>Regular Votes</h2>")
        regular_votes_html = cumulative_regular_votes.copy().reset_index().rename(columns={'index': 'Question'}).to_html(index=False, escape=False)
        f.write(regular_votes_html)

        # Integer votes table
        f.write("<h2>Integer Votes</h2>")
        integer_votes_html = cumulative_integer_votes.copy().reset_index().rename(columns={'index': 'Question'}).to_html(index=False, escape=False)
        f.write(integer_votes_html)

        # Integer votes percentages table
        f.write("<h2>Integer Votes Percentages</h2>")
        integer_votes_percentages_html = integer_votes_percentages.copy().reset_index().rename(columns={'index': 'Question'}).to_html(index=False, escape=False)
        f.write(integer_votes_percentages_html)

        # Comments table (only comment text)
        if all_comments:
            f.write("<h2>Comments</h2>")
            comments_df = pd.DataFrame(all_comments)
            if 'Comment' in comments_df.columns:
                comments_df = comments_df[['Comment']]
            comments_html = comments_df.to_html(index=False, escape=False)
            f.write(comments_html)

        # Embed the image using base64
        f.write("<h2>Votes Plot</h2>")
        f.write(f"<img src='data:image/png;base64,{encoded_img}' alt='Votes Plot'>")
        f.write("</body></html>")
    
    print(f"Aggregated results HTML saved to {html_file_path}")

# Usage
folder_path = 'surveys'
question_text = "Le formateur a communiqué de façon dynamique"
aggregate_votes_and_generate_html(folder_path, question_text)
