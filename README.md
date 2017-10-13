# Exam generator

Script generates an exam with questions and answers in random order.
List of correct answers is saved in a separate excel file.

Dependencies:
- numpy
- pandas
- docx
- and their dependencies

Script can be used for automatic default exam generation.
It requires two files: 'exam.xlsx' and 'template.docx' in the same directory as the script.

Excel file should contain one question with answers in each row. Correct answer should have a # (hash) sign at the end. Additionally, one can add a % (percent) sign at the end of every answer that should not be shuffled. Sample rows from such a file are presented below.
    
|Question Column|Answer Column|Answer Column|Answer Column|Answer Column|Answer Column| 
|:-:|:-:|:-:|:-:|:-:|:-:|
|Question 1|Answer 1|Answer 2|Answer 3|Correct Answer#|Immobile Answer%|
|Question 2|Answer 1|Answer 2|Immobile Answer%|Answer 4|Correct Answer#|
|Question 3|Correct Answer#|Answer 2|Answer 3|Immobile Answer%|Immobile Answer%|

**Important**  
Your file can't have a header. Just questions and answers.

The template word file will be populated by the script with questions. Therefore it is a very good tool to pre-shape your exam file, e.g. give it two columns, proper margins, font size and so on. If the template will have any content, questions with answers will be added at the end of the document.

In this repository you will find sample question file and a template.

### Modifying the script
This will give some additional control over shuffling and exporting process.
There is an additional function to save the exam in plain text format.
Another very useful functionality lets you generate many different versions at once.

File aesthetics is limited by my ability to control docx library (which means 'very limited') and my particular needs. You can alter the font or text intedation with ease in code. Other modifications will have to be implemented first.

Created with Python 3.6.2, Numpy 1.13.1, pandas 0.20.3, docx 0.8.6.
