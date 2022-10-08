# Genpact-exercise

Programming exercise to be done in Python and VBA.

## Programming exercise

Build a excel consolidation tool with a folder watcher.

### Basic guidelines:

- You can use either C#, python or java languages. Please follow OOP.

### Solution Description:

- Build a solution, which monitor a folder looking for new files.
- Each time a file is found, it should verify if is an excel file (.xls\* files). If is true, it should take each sheet on it and consolidate it on a master workbook file (make a copy from each sheet to the master file).
- It should have an option to choose which folder to watch.
- Every file found should be moved to 2 different folders depending if was or not a excel file
  - Processed
  - Not applicable.

### How to use

Install the dependecies and then execute the file with the command:

```sh
python -m pip install -r requirements.txt
python main.py
```

Then you are going to be prompted to insert the name of the folder that is going to be observed (the folder can exist or not).

```
Input the lookup folder ["./lookup"]:
```

3 folders should be created automatically:

- `lookup` (if you just press enter or the observed folder)
- `no_applicable`
- `processed`

And a workbook called `master_workbook.xlsx` if it doesn't exists.

