# ilias_evaluator
ILIAS Evaluator: A tool for evaluating and post-correction of tests and exams  done in the ILIAS or ILIAS E-Assessment System  

The tool is created by Silvan Rummeny. A first approach of post-correction of ILIAS tests is done  in https://github.com/P4ckP4ck/ILIAS_KlausurAuswertung /Bewerte ILIAS-Testergebnisse V1_5.py created by Eberhard Waffenschmidt, TH Köln.   

For export of the tests and exams done in ILIAS or ILIAS E-Assessment System  please use the tool https://github.com/TPanteleit/ILIAS---Test-Generator

This tool consists of: 

    - classes:
        - Test
    - general methods (in alphabetical order):
        - drop_None
        - eval_ILIAS
        - evaluate_bonus
        - evaluate_exam
        - evaluate_intermediate_tests
        - evaluate_praktika
        - import_psso_members
        - get_excel_files
        - get_originality_proof

This tool is capable of: 
    - read and process ILIAS test results
    - evaluate intermediate tests of a course regarding bonus points
    - evaluate exam of a course
    - tasks with multiple answers can be considered in evaluation
    - determine course note for each participant
    - export detailed results (e.g. for exam review for students)

This tool is limited as follows: 
    - active evaluation only for task types: Formelfrage, Single Choice
    - passive ILIAS result import only for tasks which are in ILIAS import data
TODO: 
    - feasible extension for active evaluation for task types: Multiple Choice, 
        Freitextaufgabe
