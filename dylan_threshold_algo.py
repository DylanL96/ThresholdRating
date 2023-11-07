import pandas as pd

# Load the dataset
df = pd.read_excel('cleaned_dataset.xlsx')

# Define mismatch scores based upon commonly mismatched characters and numbers
# Lower the score means it is more commeonly mistaken together, hence it is given a lower value
mismatch_scores = {
    ('B', '8'): 0.5,
    ('D', 'O'): 0.5,
    ('O', 'D'): 0.5,
    ('O', '0'): 0.5,
    ('l', '1'): 0.5,
    ('I', '1'): 0.5,
    ('S', '5'): 0.5,
    ('G', '6'): 0.5,
    ('g', '9'): 0.5,
    ('Z', '2'): 0.5,
    ('q', '9'): 0.5,
    ('E', 'F'): 0.75,
    ('M', 'N'): 0.75,
    ('Q', 'O'): 0.75,
    ('U', 'V'): 0.75,
    
}
# Set constants for scoring
# This was based on the fuzzy logic document
score_per_letter = 0.15
max_letter_score = 1.25
min_letter_score = 0.60

# Function to calculate mismatch score and determine match for each row
def calculate_mismatch_score_and_match(row):
    permit = str(row['permit'])
    scanned = str(row['scanned'])
    mismatch_score = 0
    
    # Calculate mismatch score
    # Compares the letters from the permit and scanned column in tuple fashion
    for permit_char, scanned_char in zip(permit, scanned):
        if permit_char != scanned_char:
            # Get the mismatch score from the dictionary, default to 2.0 if not found
            # mismatch_score is the dictionary where each key is a tuple of two characters that are commonly mistaken
            # value is an arbitrary score of how much of a mismatch those two characters are
            # If the tiple doesnt exist in the dictionary, I assume that the two characters being compared are not a common
            # mistaken, and so if the mistake was made, then the value would be 2.0
            mismatch_score += mismatch_scores.get((permit_char, scanned_char), 2.0)
    
    # Calculate threshold based on the length of the permit
    threshold = len(permit) * score_per_letter

    # Enforce max and min scores per character
    threshold = min(threshold, max_letter_score * len(permit))
    if mismatch_score < min_letter_score:
        threshold = 0  # Require exact match for mismatch scores below the min
    
    # Determine match and return both mismatch_score and match
    match = mismatch_score <= threshold
    return pd.Series([mismatch_score, threshold, match], index=['mismatch_score', 'threshold', 'match'])

# Apply the function and create new columns for mismatch_score and match
df[['mismatch_score', 'threshold', 'match']] = df.apply(calculate_mismatch_score_and_match, axis=1)

# Save the updated DataFrame to a new Excel file
df.to_excel('updated_cleaned_dataset.xlsx', index=False)