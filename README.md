# Automatically Create and Evaluate Excel-Excercises
We are doing a PC Course focussing on Excel and Word. Due to resource limitations, the exercises have to be created and evaluated automatically. Each student
will get an inidividual exercise, i.e. the principal exercises per student are all the same, however the specific numbers, sizes of tables are different. The
exercise is given as a pdf, each student gets his excel-file, i.e. the file named according to his student-ID (Matrikelnummer).
When the results are all uploaded to Ilias (our Learn Management System), the solutions are uploaded in a specific directory structure. The check_exercises
script evaluates the exercises and documents the results into an excel-file.

The scripts to create and check the excel-files make use of the openpyxl-library.
