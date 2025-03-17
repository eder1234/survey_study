# Adjusted extraction function considering explicit locations and column labels
def extract_votes(df, start_row, start_col, num_questions=18, classes_per_question=5, row_step=3):
    questions = []
    votes_data = []
    
    for q in range(num_questions):
        question_row = start_row + q * row_step
        votes_row = question_row + 1

        # Extract question text
        question_text = df.iloc[question_row, start_col]
        questions.append(question_text)

        # Extract votes per class
        votes = df.iloc[votes_row, start_col+1:start_col+1+classes_per_question].values
        votes_data.append(votes)

    # Define class labels explicitly
    class_labels = ['Tout à fait d\'accord', 'D\'accord', 'Indécis', 'En désaccord', 'Tout à fait en désaccord']
    votes_df = pd.DataFrame(votes_data, columns=class_labels, index=questions)
    
    return votes_df

# Adjust indices according to explicit verification
start_row, start_col = 72, 1  # As explicitly found above

# Extract votes dataframes from both files (assuming the second file has a similar structure)
votes_df1 = extract_votes(df1, start_row, start_col)
votes_df2 = extract_votes(df2, start_row, start_col)

# Display results
import ace_tools as tools
tools.display_dataframe_to_user(name="Votes Evaluation UE1 - Santé publique", dataframe=votes_df1)
tools.display_dataframe_to_user(name="Votes Evaluation UE2 - SHS", dataframe=votes_df2)

votes_df1.head(), votes_df2.head()

