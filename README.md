# RDVBA Project Utils

[RDVBA] is a VBA IDE extension that provides several developing tools missing from the stock code editor. Hierarchy-based project representation within the IDE is one such prominent tool. Also, it can synchronize the entire project between the host application file and an independent folder, enabling version management. However, RDVBA cannot synchronize the virtual project hierarchy with the project folder.

Virtual and physical hierarchies synchronization is instrumental for developing packages/libraries and transferring them between projects. Such a feature would also make it simpler to quickly check the code in an exported module, which may be more convenient than opening the original application file. RDVBA Project Utils uses pure VBA to implement such hierarchy synchronization functionality, including individual project subfolder synchronization.

## Code description

### Usage

RDVBA Project Utils synchronizes all code modules between the ActiveProject (from the ActiveWorkbook) and a folder called "Project" located next to the Excel file (and it may work for Word and Access as well). The "Project\Common\Project Utils" contains the ProjectUtils class, implementing functionality, and the ProjectUtilsSnippets module with code snippets, running the tasks.

- Import/Export library references (Tools->References):
  *ProjectUtils.ReferencesSaveToFile* and
  *ProjectUtils.ReferencesAddFromFile*.
- Import/Export project modules/structure:
  *ProjectUtils.ProjectFilesExport* and
  *ProjectUtils.ProjectFilesImport*.

*ProjectUtils.ProjectFilesExport* takes one optional argument - folder prefix to be exported relative to the "Project" folder. If not provided, the entire project is exported.

*ProjectUtils.ProjectFilesImport* takes two optional arguments: a folder prefix to be imported relative to the "Project" folder and a Boolean flag indicating whether to skip importing files from the top imported directory. The default behavior: import the entire "Project" folder and skip files that are immediate children of the "Project" folder.

For demo snippets, look for a function in the *ProjectUtilsSnippets* module matching the method name from the *ProjectUtils* class.

### References

*ProjectUtils.ReferencesSaveToFile* saves the details of activated references (Tools->References) to a tab-separated file ("References.xsv" in the "Project" directory), with each record containing Name, GUID, Major, Minor, and FullPath fields.
*ProjectUtils.ReferencesAddFromFile* parses and activates previously saved references from a file.
This functionality can be used for the automatic activation of references required by a package during its import.

### Export

*ProjectStructureParse* goes through all code modules (VBComponents) in the ActiveProject, reads "'@Folder" annotation for each component, and assembles this information for project export. Modules, which do not have "'@Folder" annotation are placed in one of four subfolders in the "Common" folder selected based on their type.
*ProjectStructureExport* creates folder structure in the Project folder reproducing that of virtual folders.
*ProjectFilesExport* exports the code modules.

### Import

*WalkTree*/*WalkTreeCore* collects folder information in the Project folder.
*CollectFiles* gathers file information.
*ProjectFilesImport* runs actual import, overwriting existing modules, and if "References.xsv" is found in the root, it will be applied as well.

## Limitations

At present, there are no tests in the project (there is a QSort test module, which is not related to the core functionality). I have interactively tested (running the snippets as indicated above) and verified functionality manually in Excel 2002/XP only, so reasonable caution should be exercised, especially with the import process, which overwrites any modules with the same name.

The import procedure assumes that the file names match the name definition inside modules, and it deletes existing modules with conflicting names. Document modules cannot be created/removed via the VBComponents collection, and any such attempt causes an error, which is silently ignored. The import of a document module results in a standard class module being created and automatically renamed. The code from this module replaces the code-behind in the corresponding document module, and the “temporary” class module is then removed.

One more feature, which is probably necessary but not yet implemented, is the addition/updating of the "'@Folder" annotation in the imported modules.

[RDVBA]: https://github.com/rubberduck-vba
