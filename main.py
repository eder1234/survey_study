import pandas as pd
import glob
import os
import matplotlib
matplotlib.use('Agg')  
import matplotlib.pyplot as plt

# Function to find question start position
def find_question_start(df, question_text):
    location = df.isin([question_text])
    result = location.stack().idxmax()
    row, col_name = result
    col = df.columns.get_loc(col_name)  # Convert column name to integer
    return row, col

# Extraction function for votes
def extract_votes(df, start_row, start_col, num_questions=18, classes_per_question=5, row_step=3):
    votes_data = []
    questions = []
    
    for q in range(num_questions):
        question_row = start_row + q * row_step
        votes_row = question_row + 1

        question_text = df.iloc[question_row, start_col]
        votes = df.iloc[votes_row, start_col+1:start_col+1+classes_per_question].values.astype(float)

        questions.append(question_text)
        votes_data.append(votes)

    class_labels = ['Tout à fait d\'accord', 'D\'accord', 'Indécis', 'En désaccord', 'Tout à fait en désaccord']
    return pd.DataFrame(votes_data, columns=class_labels, index=questions)

# Main aggregation function
def aggregate_votes_from_folder(folder_path, question_text):
    all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

    cumulative_votes = None

    for file in all_files:
        df = pd.read_excel(file)
        start_row, start_col = find_question_start(df, question_text)
        file_votes = extract_votes(df, start_row, start_col)

        if cumulative_votes is None:
            cumulative_votes = file_votes
        else:
            cumulative_votes += file_votes

    file_path = os.path.join(folder_path, 'aggregated_votes.csv')
    cumulative_votes.to_csv(file_path, encoding='utf-8-sig')

    print(f"Aggregated votes successfully saved to {file_path}")
    df = pd.read_csv(file_path)

    # Rename the columns for clarity
    df.rename(columns={'Unnamed: 0': 'Question'}, inplace=True)

    # Compute percentages
    vote_columns = ['Tout à fait d\'accord', 'D\'accord', 'Indécis', 'En désaccord', 'Tout à fait en désaccord']
    df_percentages = df.copy()
    df_percentages[vote_columns] = df[vote_columns].div(df[vote_columns].sum(axis=1), axis=0) * 100

    # Plotting
    fig, ax = plt.subplots(figsize=(12, 8))

    df_percentages.set_index('Question')[vote_columns].iloc[::-1].plot(kind='barh', stacked=True, ax=ax, colormap='viridis')

    plt.xlabel('Percentage (%)')
    plt.ylabel('Question')
    plt.title('Percentage of Votes per Class')
    plt.legend(title='Vote Type', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()

    plt.savefig("results.png")

# Usage example (replace with your folder path)
folder_path = 'surveys'
question_text = "Le formateur a communiqué de façon dynamique"

aggregate_votes_from_folder(folder_path, question_text)