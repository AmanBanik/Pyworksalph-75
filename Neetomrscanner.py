import pandas as pd
import numpy as np 

# Read Excel files
try:
    # Read as object to prevent premature conversion.
    # We expect '0' for unattempted questions from the user's manual modification.
    response_df = pd.read_excel(r"C:\PY.programes\neetrohit\neet_response.xlsx", index_col="Question", dtype={"Answer": object})
    response = response_df["Answer"]
except FileNotFoundError:
    print("Error: 'neet_response.xlsx' not found. Please check the path.")
    exit()
except KeyError:
    print("Error: 'Answer' column not found in 'neet_response.xlsx'.")
    exit()

try:
    # Key should still be processed robustly for its own content.
    key_df = pd.read_excel(r"C:\PY.programes\neetrohit\answer_key.xlsx", index_col="Question", dtype={"Answer": object})
    key = key_df["Answer"]
except FileNotFoundError:
    print("Error: 'answer_key.xlsx' not found. Please check the path.")
    exit()
except KeyError:
    print("Error: 'Answer' column not found in 'answer_key.xlsx'.")
    exit()

# Convert each answer in the key to a clean set of integer options
def parse_key_answer_to_int(answer):
    try:
        # Key answers should ideally not be blank, but handle defensively
        if pd.isna(answer) or (isinstance(answer, str) and str(answer).strip() == ""):
            return set()
        
        options_as_str = str(answer).replace(" ", "").split(',')
        
        int_options = set()
        for opt_str in options_as_str:
            try:
                int_options.add(int(float(opt_str)))
            except ValueError:
                # If a part of the key cannot be converted, it's ignored.
                print(f"Warning: Key answer part '{opt_str}' for question could not be converted to integer.")
                pass
        return int_options
    except Exception as e:
        print(f"Error parsing key answer (key file): {answer} -> {e}")
        return set()

correct_sets = key.apply(parse_key_answer_to_int)

# Normalize user responses to integer values
def normalize_response_to_int(resp):
    try:
        # If resp is NaN (unlikely if blanks are replaced with 0) or purely empty string,
        # it indicates an unexpected blank. Treat it as None to be safe, though 0 is expected.
        if pd.isna(resp) or (isinstance(resp, str) and resp.strip() == ""):
            return None # Should ideally not be reached if 0s are filled in

        resp_str_cleaned = str(resp).strip()
        
        try:
            val = int(float(resp_str_cleaned))
            # If the value is 0, this is the designated 'unattempted' value
            if val == 0:
                return 0
            # Otherwise, return the parsed integer
            return val
        except ValueError:
            # If it's not 0 and not a valid number, treat as an unparseable/invalid attempt (0 marks)
            print(f"Warning: User response '{resp_str_cleaned}' could not be converted to integer (and is not 0). Treating as unattempted.")
            return None # Represents invalid/unattempted (0 marks)
            
    except Exception as e:
        print(f"Error normalizing response (response file): {resp} -> {e}")
        return None # General error also leads to unattempted

response = response.apply(normalize_response_to_int)

# Debugging: Print a sample of processed data
print("--- Sample of Processed Data ---")
print("Response (first 10, checking for 0 for unattempted):")
print(response.head(10))
print("\nCorrect Sets (first 5):")
print(correct_sets.head())
print("------------------------------\n")

# Initialize counters for summary
num_correct = 0
num_incorrect = 0
num_unattempted = 0
total_score = 0

detailed_results = [] # To store details for debugging

for question in response.index:
    # Get the normalized answer. It will be an integer (0, 1, 2, 3, 4) or None for truly unparseable data.
    given = response.get(question, None) 
    correct = correct_sets.get(question, set())

    status = ""
    points = 0

    # Logic based on new strategy: 0 means unattempted (0 marks)
    if given is None: # This handles cases where normalize_response_to_int returned None (e.g., non-numeric and not 0)
        status = "Unattempted (Invalid/Non-numeric)"
        points = 0
        num_unattempted += 1
    elif given == 0: # Explicitly checking for 0 as the unattempted marker
        status = "Unattempted"
        points = 0
        num_unattempted += 1
    elif given in correct: # Correct answer
        total_score += 4
        points = 4
        status = "Correct"
        num_correct += 1
    else: # Answered, but incorrect
        total_score -= 1
        points = -1
        status = "Incorrect"
        num_incorrect += 1

    correct_display = ', '.join(map(str, sorted(list(correct)))) if correct else "N/A"
    
    # Adjusting "Your Answer" display if it's 0 (meaning unattempted)
    display_answer = "UNATTEMPTED" if given == 0 else (given if given is not None else "UNATTEMPTED (INVALID)")

    detailed_results.append({
        "Question": question,
        "Your Answer": display_answer,
        "Correct Answer(s)": correct_display,
        "Status": status,
        "Points": points,
        "Running Total": total_score
    })

# Print detailed results for debugging
print("\n--- Detailed Scoring Results ---")
for result in detailed_results:
    print(f"Q{result['Question']}: Your Answer='{result['Your Answer']}', Correct='{result['Correct Answer(s)']}', Status='{result['Status']}', Points={result['Points']}, Running Total={result['Running Total']}")

print("\n--- Summary ---")
print(f"Total Questions: {len(response.index)}")
print(f"Correct Answers: {num_correct}")
print(f"Incorrect Answers: {num_incorrect}")
print(f"Unattempted Questions: {num_unattempted}")
print(f"\n🎯 Final Score: {total_score}")

# Optional: Save detailed results to a DataFrame for easier analysis
results_df = pd.DataFrame(detailed_results)
print("\n--- Detailed Results DataFrame ---")
print(results_df)
results_df.to_excel("neet_scoring_details.xlsx", index=False)
print("\nDetailed results saved to 'neet_scoring_details.xlsx'")