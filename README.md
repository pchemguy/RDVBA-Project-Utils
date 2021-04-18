# RDVBA Project Utils

[RDVBA](https://github.com/rubberduck-vba) greatly improves the VBA programming experience in several ways within the stock IDE. Its Code Browser with virtual folders is one such great feature. It also permits exporting/importing all code modules at once. However, the virtual folder structure is not created on the drive.

While I mostly edit VBA code from within the RDVBA enhanced IDE, I often need to check things in other VBA projects to copy a piece of code or a set of modules. I do not open such projects in the IDE but rather go to the project folder using a file manager. The lack of project structure on the hard drive makes it difficult to work with the project this way. Additionally, I have some modules implementing common functionality, which are therefore used in different projects. I am not aware of a way of importing/exporting such sets straightforwardly. For these reasons, it would be great if RDVBA exported the virtual folders as well and imported files from a structured project.

Code in this project implements such functionality in VBA with the hope that at some point it such a feature will be integrated into RDVBA. 

## Code description

This project assumes the convention that all code modules are exported to (with folder structure) and can be imported from the "PROJECT" folder sitting next to the Excel file (and it may work for Word and Access as well).

The "Project\Common\Project Utils" contains the core code, namely, the `ProjectUtils` class, implementing functionality, and `ProjectUtilsSnippets` with code snippets, running the tasks. Two modules inside the "Project\Common\QuickSort" folder implement the basic "QuickSort" algorithm and are not used by the project.

### Usage

Importantly, at present, ActiveProject is set as the target for import/export. That means, that the target application file must be the active one (e.g. ActiveWorkbook). Import/export is performed to/from the "PROJECT" folder or its subfolders. Arbitrary locations are not supported.

- Import/Export library references (Tools->References) - run `ProjectUtilsSnippets.ReferencesSaveToFile` and `ProjectUtilsSnippets.ReferencesAddFromFile`.
- Import/Export project modules/structure - run `ProjectUtilsSnippets.ProjectFilesExport` and `ProjectUtilsSnippets.ProjectFilesImport`.

`ProjectUtils.ProjectFilesExport` takes one optional argument - folder prefix to be exported relative to the "PROJECT" folder. If not provided, the entire project is exported.
`ProjectUtils.ProjectFilesImport` takes two additional arguments: a folder prefix to be imported relative to the "PROJECT" folder (if not provided, the entire project is imported) and a Boolean flag indicating whether to skip importing files from the top imported directory (by default is True to skip, when imported entire project "PROJECT" folder and False, when a subfolder is indicated as the first argument.

### References

`ReferencesSaveToFile` saves the details of activated references (Tools->References) into a tab-separated table (file "References.xsv" in the "Project" directory), with each record containing "Name", "GUID", "Major", "Minor", and "FullPath" fields. `ReferencesAddFromFile`, in turn, reads and parses this file and activates all references. The idea is that common functionality can be implemented as "packages". If such a package requires certain references, a references file can be included in its top folder, and it will be parsed and applied during the package import process. `ProjectUtilsSnippets.ReferencesSaveToFile` and `ProjectUtilsSnippets.ReferencesAddFromFile` can be used to run this routines.

### Export

`ProjectStructureParse` goes through all code modules (VBComponents) in the ActiveProject, reads "'@Folder" annotation for each component, and assembles this information for project export. Modules, which do not have "'@Folder" annotation are placed in one of four subfolders in the "Common" folder selected based on their type (demo run with `ProjectUtilsSnippets.ProjectStructureParse`).

`ProjectStructureExport` creates folder structure in the Project folder reproducing the structure of virtual folders (demo run with `ProjectUtilsSnippets.ProjectStructureExport`).

 `ProjectFilesExport` exports the code modules  (demo run with `ProjectUtilsSnippets.ProjectFilesExport`).

### Import

`WalkTree`/`WalkTreeCore` collects folder information in the Project folder (`ProjectUtilsSnippets.WalkTree`).

`CollectFiles` gathers file information (`ProjectUtilsSnippets.CollectFiles`).

`ProjectFilesImport` runs actual import, overwriting existing modules, and if "References.xsv" is found in the root, it will be applied as well (`ProjectUtilsSnippets.ProjectFilesImport`).

## Limitations

At present, there are no tests in the project (there is a QSort test module, which is not related to the core functionality). I have interactively tested (running the snippets as indicated above) and verified functionality manually on Excel 2002/XP only, so reasonable caution should be exercised, especially with the import process, which overwrites any modules with the same name.

The import process assumes that the file names match the name definition inside modules (no checks performed) and if a VBComponent with the same name present in the project, it is removed before import. The only exception to this is for documents modules. While the component removal code does not check the type of the module, an attempt to delete such a component via the VBComponents collection is rejected with an error as expected. Such an error is ignored, and during the actual import step, the routine checks the type of the module. For document modules, import results in a standard class module being created and automatically renamed. The code from such a module is copied, and it replaces the code-behind in the corresponding document module. The "temporary" class module is then removed from the project.

One more feature, which is probably necessary, but not yet implemented, is the addition/updating of the `'@Folder` annotation in the imported modules. It can be added straightforwardly, and I might add it at some point.
