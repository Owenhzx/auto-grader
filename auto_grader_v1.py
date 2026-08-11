"""
Automated Multiple-Choice Grading System — v1

Batch-grades multiple-choice exams from a folder of per-student Excel
answer files: parses a flexible answer-key syntax (single- and
multiple-answer questions), normalises student responses, computes
scores, and exports per-student results plus cohort-level question
analytics to a multi-sheet Excel report.

Note: this is the English-localised version of the original script;
the tool was originally deployed with Chinese-language prompts and
generated comments.
"""

import os
import pandas as pd
from collections import defaultdict


def parse_answer_str(answer_str):
    """
    Parse the answer-key string into an answer dictionary.
    :param answer_str: the answer-key string
    :return: (answer_key, total_questions) — answer dictionary and question count
    """
    # Check whether the string uses the slash delimiter
    if '/' in answer_str:
        # Slash-delimited format (may contain multiple-answer questions)
        answer_parts = answer_str.split('/')
        answer_key = {}
        for i, part in enumerate(answer_parts, start=1):
            # Remove all spaces (visual separators only)
            cleaned = part.replace(" ", "").upper()
            if cleaned:
                # Deduplicate with a set and sort into canonical order
                unique_chars = sorted(set(cleaned))
                sorted_answer = ''.join(unique_chars)
                answer_key[f'Q{i}'] = sorted_answer
            else:
                answer_key[f'Q{i}'] = ""
        total_questions = len(answer_parts)
    else:
        # No slashes: every question is single-answer.
        # Split into space-separated groups; each character is one question.
        groups = answer_str.split()
        answer_key = {}
        question_count = 1

        for group in groups:
            # Remove spaces and convert to upper case
            cleaned_group = group.replace(" ", "").upper()
            # Each character in the group represents one single-answer question
            for char in cleaned_group:
                answer_key[f'Q{question_count}'] = char
                question_count += 1

        total_questions = question_count - 1

    return answer_key, total_questions


def clean_answer(ans):
    """
    Clean and normalise a student's answer.
    :param ans: the raw student answer
    :return: the normalised answer string
    """
    if pd.isna(ans):
        return ""

    # Convert to string and upper case
    ans_str = str(ans).upper()

    # Remove all spaces (visual separators only)
    cleaned = ans_str.replace(" ", "")

    # Deduplicate and sort (for multiple-answer questions)
    if cleaned:
        unique_chars = sorted(set(cleaned))
        cleaned = ''.join(unique_chars)

    return cleaned


def process_student_file(file_path, answer_key, total_questions, error_counter):
    """
    Grade a single student's answer file.
    :param file_path: path to the student's answer file
    :param answer_key: the answer dictionary
    :param total_questions: total number of questions
    :param error_counter: per-question error counter
    :return: a dictionary with the student's results
    """
    try:
        # Use the file name (without extension) as the student ID
        student_id = os.path.splitext(os.path.basename(file_path))[0]

        # Read the student's answers
        df = pd.read_excel(file_path, header=None)

        # Extract the first column of answers as a list
        answers = df.iloc[:, 0].tolist()

        # Normalise every answer
        processed_answers = []
        for i in range(total_questions):  # Only process answers within the question count
            if i < len(answers):
                cleaned = clean_answer(answers[i])
                processed_answers.append(cleaned)
            else:
                # Missing answers are treated as empty strings
                processed_answers.append("")

        # Compute the score and collect wrong answers
        score = 0
        wrong_questions = []

        for i in range(total_questions):
            q = f'Q{i+1}'
            student_ans = processed_answers[i]
            correct_ans = answer_key[q]

            # Compare answers (exact match required)
            if student_ans == correct_ans:
                score += 1
            else:
                wrong_questions.append(q)
                # Record the error for this question
                error_counter[q] += 1

        # Generate the written comment
        pass_rate = score / total_questions
        comment = (f"This morning's assessment was the day-2 quiz with {total_questions} "
                   f"multiple-choice questions. {student_id} scored {score}/{total_questions}.")
        if pass_rate > 0.5:
            comment += " Excellent work — keep it up!"
        else:
            comment += " Keep working at it!"

        # Return the result
        return {
            'Student_ID': student_id,
            'Total_Score': score,
            'Wrong_Questions': ",".join(wrong_questions),
            'Comment': comment
        }

    except Exception as e:
        print(f"Error while processing file {file_path}: {e}")
        return None


def grade_exams_from_folder(input_folder, output_path, answer_str):
    """
    Batch-grade all student answer files in a folder.
    :param input_folder: folder containing the student answer files
    :param output_path: path for the results workbook
    :param answer_str: the answer-key string
    """
    try:
        # Parse the answer-key string
        answer_key, total_questions = parse_answer_str(answer_str)
        if total_questions == 0:
            print("Error: the answer key is empty — nothing to grade")
            return
        print(f"Answer key parsed: {total_questions} questions in total")
        print(f"Answer key: {answer_key}")

        # Collect all student answer files
        student_files = []
        for file in os.listdir(input_folder):
            if file.endswith(('.xlsx', '.xls')):
                student_files.append(os.path.join(input_folder, file))

        if not student_files:
            print(f"No Excel files found in folder {input_folder}")
            return

        print(f"Found {len(student_files)} student answer files")

        # Initialise the per-question error counter
        error_counter = defaultdict(int)

        # Process every student file
        results = []
        for file_path in student_files:
            result = process_student_file(file_path, answer_key, total_questions, error_counter)
            if result:
                results.append(result)

        if not results:
            print("No student files could be graded successfully")
            return

        # Build the results DataFrame
        result_df = pd.DataFrame(results)

        # Ensure a consistent column order
        result_df = result_df[['Student_ID', 'Total_Score', 'Wrong_Questions', 'Comment']]

        # Flag questions for in-class review (wrong for more than 1/3 of the cohort)
        total_students = len(results)
        review_threshold = total_students / 3
        review_questions = []

        # Build the question-analysis DataFrame
        question_analysis_df = []
        for i in range(1, total_questions + 1):
            q = f'Q{i}'
            error_count = error_counter[q]
            error_rate = error_count / total_students if total_students > 0 else 0
            needs_review = error_count > review_threshold

            question_analysis_df.append({
                'Question': q,
                'Errors': error_count,
                'Error Rate': f"{error_rate:.2%}",
                'Needs Review': 'Yes' if needs_review else 'No'
            })

            if needs_review:
                review_questions.append(q)

        analysis_df = pd.DataFrame(question_analysis_df)

        # Save the results to Excel (multiple sheets)
        with pd.ExcelWriter(output_path) as writer:
            result_df.to_excel(writer, sheet_name='Student Scores', index=False)
            analysis_df.to_excel(writer, sheet_name='Question Analysis', index=False)

            # Add the review-suggestions sheet
            if review_questions:
                review_df = pd.DataFrame({
                    'Question': review_questions,
                    'Errors': [error_counter[q] for q in review_questions],
                    'Error Rate': [f"{error_counter[q]/total_students:.2%}" for q in review_questions]
                })
                review_df.to_excel(writer, sheet_name='Review Suggestions', index=False)

        print(f"Grading complete! Results saved to: {output_path}")
        print(f"Processed {total_students} students in total")

        # Print the review suggestions
        if review_questions:
            print("\nQuestions recommended for in-class review:")
            for q in review_questions:
                print(f"- {q}: {error_counter[q]}/{total_students} wrong "
                      f"({error_counter[q]/total_students:.2%})")
        else:
            print("\nNo questions require special review")

    except Exception as e:
        print(f"Error: {e}")


def get_user_input():
    """Collect all required input from the user."""
    print("=" * 50)
    print("Automated Multiple-Choice Grading System")
    print("=" * 50)

    # Get the answer-key string
    print("\n[Step 1/3] Enter the answer key:")
    # print("Format notes:")
    # print("- Single-answer questions:")
    # print("  1-7 questions: enter consecutive characters, e.g. 'ABCDABA'")
    # print("  8+ questions: add a space every 5 questions, e.g. 'ABCDE FGHIJ KLMNO'")
    # print("- Multiple-answer questions: separate questions with slashes, e.g. 'AB/A/DE/ABC/DD'")
    # print("  Spaces are allowed inside slash-delimited questions, e.g. 'AB/A/DE/ABC/DD B/C/D/E/DE DA/A'")
    answer_str = input("Answer-key string: ").strip()

    # Get the input folder
    print("\n[Step 2/3] Enter the folder containing student answer files:")
    print(r"Example: C:\Users\me\Desktop\quiz-answers")
    input_folder = input("Folder path: ").strip()

    # Validate that the folder exists
    while not os.path.exists(input_folder):
        print(f"Error: folder '{input_folder}' does not exist")
        input_folder = input("Enter a valid folder path: ").strip()

    # Get the output file path
    print("\n[Step 3/3] Enter the path for the results workbook:")
    print(r"Example: C:\Users\me\Desktop\grading-results.xlsx")
    output_path = input("Output file path: ").strip()

    # Ensure the output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    return answer_str, input_folder, output_path


if __name__ == "__main__":
    # Collect user input
    answer_str, input_folder, output_path = get_user_input()

    # Run the grading
    grade_exams_from_folder(input_folder, output_path, answer_str)

    # Exit prompt
    print("\nDone! Press Enter to exit...")
