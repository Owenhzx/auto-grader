# Automated Multiple-Choice Grading System

A Python tool that batch-grades multiple-choice exams from a folder of
per-student Excel files, exporting scores, per-student comments, and
cohort-level question analytics to a single workbook.

Built for a teaching academy, where it processed 300+ submissions per day and
saved each teacher roughly 1.5 hours of marking daily.

**Python · Jul 2025**

## Features

- **Flexible answer-key syntax** — plain strings for single-answer papers
  (`ABCDE FGHIJ`), slash-delimited for papers mixing single- and
  multiple-answer questions (`AB/A/DE/ABC`).
- **Answer normalisation** before comparison — case, whitespace, duplicates,
  and ordering — so `"b a"`, `"AB"`, and `"BA"` all match a key of `AB`.
- **Automatic review flagging** for any question answered incorrectly by more
  than a third of the cohort.
- **Knowledge-point tagging (v2)** — maps wrong answers to their topics and
  generates a personalised written comment for every student.

## Usage

```bash
pip install pandas openpyxl
python auto_grader_v2.py
```

The script prompts for the answer key, per-question knowledge points, the
folder of student files, and the output path. Student files should be
`.xlsx`/`.xls` with answers in the first column; the file name is used as the
student ID.

## Output

A single workbook with three sheets:

| Sheet | Contents |
|---|---|
| Student Scores | Student ID, total score, wrong questions, written comment |
| Question Analysis | Per-question error count, error rate, review flag, topic |
| Review Suggestions | Only the questions that crossed the review threshold |

## Versions

`auto_grader_v1.py` (270 lines) — core grading and cohort analytics
`auto_grader_v2.py` (327 lines) — adds knowledge-point tagging and comments

