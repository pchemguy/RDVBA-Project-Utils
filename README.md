# RDVBA Project Utils

[RDVBA](https://github.com/rubberduck-vba) greatly improves VBA programming experience in number of ways within the stock IDE. Its Code Browser with virtual folders is one such great feature. It also permits exporting/importing all code modules at once. However, the virtual folder structure is not created on the drive.

While I mostly edit VBA code from within the RDVBA enhanced IDE, I often need to check things in other VBA projects to copy a piece of code or a set of modules. I do not open such projects in the IDE, but rather go to the project folder using a file manager. The lack of project structure on the hard drive makes it difficult to work with project in such a way. Additionally, I have some modules implementing common functionality, which are therefore used in different projects. I am not aware of a way of importing/exporting such sets straightforwardly. For this reasons, it would be great if RDVBA exported the virtual folders as well and imported files from a structured project.

Code in this project implements such a functionality in VBA with the hope that at some point it such a feature will be integrated into RDVBA. 

## Overview

This project assumes the convention that all code modules are exported to (with folder structure) and can be imported from "Project" folder sitting next to the Excel file (and it may work for Word and Access as well).

This project itself contains four modules inside the "Project\Common" folder.

Two modules inside the `QuickSort` folder implement basic "QuickSort" algorithm and are currently not used by the project.

The "Project\Common\Project Utils" contains the actual code, namely, `ProjectUtils` class, implementing functionality and `ProjectUtilsSnippets` with code snippets, running the tasks.

Importantly, at present, ActiveProject is set as the target for import/export. That means, that the target application file must be the active one (e.g. ActiveWorkbook).

### References

`ReferencesSaveToFile` saves the details of activated references (Tools->References) into a tab separated table (file "References.xsv" in the "Project" directory), with each record containing "Name", "GUID", "Major", "Minor", and "FullPath" fields. `ReferencesAddFromFile`, in turn, reads and parses this file and activates all references. The idea is that common functionality can be implemented as "packages". If such a package requires certain references, a references file can be included in its top folder, and it will be parsed and applied during the package import process. `ProjectUtilsSnippets.ReferencesSaveToFile` and `ProjectUtilsSnippets.ReferencesAddFromFile` can be used to run this routines.

### Export

`ProjectStructureParse` goes through all code modules (VBComponents) in the ActiveProject, reads "'@Folder" annotation for each component, and assembles this information for project export. Modules, which do not have "'@Folder" annotation are placed in one of four subfolders in the "Common" folder selected based on their type (demo run with `ProjectUtilsSnippets.ProjectStructureParse`).

`ProjectStructureExport` creates folder structure in the Project folder reproducing the structure of virtual folders (demo run with `ProjectUtilsSnippets.ProjectStructureExport`).

 `ProjectFilesExport` exports the code modules  (demo run with `ProjectUtilsSnippets.ProjectFilesExport`).

### Import

`WalkTree`/`WalkTreeCore` collects folder information in the Project folder (`ProjectUtilsSnippets.WalkTree`).

`CollectFiles` gathers file information (`ProjectUtilsSnippets.CollectFiles`).

`ImportFiles` runs actual import, overwriting existing modules, and if "References.xsv" is found in the root, it will be applied as well (`ProjectUtilsSnippets.ImportFiles`).