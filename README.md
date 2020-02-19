# MENTOR/MENTEE SORTER
This is a program I created for the Student Affairs and Engagement (SAE) team at the University of Sydney’s Faculty of Arts and Social Sciences. Specifically, this program is designed to assist in the management of the Faculty of Arts and Social Sciences’ Mentorship Program by automating the process of matching mentors with mentees in preparation for the running of the Mentorship Program each semester.
This program was written in Visual Basic for Applications (VBA) and is functional within the ‘Microsoft Visual Basic for Applications’ environment.

## Download:
A working version of this program and a set of sample student data can be downloaded here: https://drive.google.com/drive/folders/1KFkRDp1szwqjV_fdRwB62Df6mnOpalR9?usp=sharing.

## In order for this project to function, the following datasets are needed:
- Degree-by-Group list;
- Mentor/Mentee list

## 1. 'Degree-by-Group' list
All degrees are assigned to a group (e.g. the Bachelor of Economics belongs to group '1', the Bachelor of Arts belongs to group '3', etc.) There is a total of 3 groups - '1', '2' and '3' - and a roughly 80 unique degrees in each cohort (although this number fluctuates each semester). There is a one-to-many relationship between degree and group.

For this program to function, it must be able to reference a dataset that lists degree names and their corresponding group.
Example:

-	| DegreeName | Group |
-	| Bachelor of Visual Arts | 1 |
-	| Bachelor of Visual Arts/Bachelor of Advanced Studies | 1 |
-	| Bachelor of Visual Arts/Bachelor of Advanced Studies | 1 |
-	| Bachelor of Education | 2 |
-	| Bachelor of Education (Early Childhood) | 2 |
-	| Bachelor of Arts | 3 |
-	| Bachelor of Arts (Languages) | 3 |

## 2. Mentor/Mentee list
In order for this program to accurately convert mentor and mentee data into instances of classes, columns must be ordered in the following configuration:

---Mentor list---
- Column 1 (A): SID
- Column 2 (B): Program
- Column 3 (C): FirstName
- Column 4 (D): LastName
- Column 5 (E): Course
- Column 6 (F): Major 1
- Column 7 (G): Major 2
- Column 8 (H): Email
- Column 9 (I): Student Type
- Column 10 (K): Int
- Column 11 (L): DL
- Column 12 (M): 25+

---Mentee list---
- Column 1 (A): SID
- Column 2 (B): FirstName
- Column 3 (C): LastName
- Column 4 (D): Intl
- Column 5 (E): Dalyell
- Column 6 (F): 25+
- Column 7 (G): Course
- Column 8 (H): Major 1
- Column 9 (I): Major 2
- Column 10 (K): Email
- Column 11 (L): Dietary
- Column 12 (M): Mobility

This complied with the configuration of columns in the raw data that is exported by the university’s registration system at the time of this program’s production.

## Notes:
- 'Degree' and 'Course' are used interchangeably in this project. It is common parlance to refer to a student’s degree as their course, and vice-versa. In this context degree == course.
