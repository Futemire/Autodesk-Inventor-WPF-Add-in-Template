Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices
Imports System.Windows

' General Information about an assembly is controlled through the following
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("$addinname$")>
<Assembly: AssemblyDescription("$addindescription$")>
<Assembly: AssemblyCompany("$companyname$")>
<Assembly: AssemblyProduct("$addinname$")>
<Assembly: AssemblyCopyright("Copyright © $companyname$ $year$")>
<Assembly: AssemblyTrademark("")>
<Assembly: ComVisible(False)>

'In order to begin building localizable applications, set
'<UICulture>CultureYouAreCodingWith</UICulture> in your .vbproj file
'inside a <PropertyGroup>.  For example, if you are using US English
'in your source files, set the <UICulture> to "en-US".  Then uncomment the
'NeutralResourceLanguage attribute below.  Update the "en-US" in the line
'below to match the UICulture setting in the project file.

'<Assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.Satellite)>

'The ThemeInfo attribute describes where any theme specific and generic resource dictionaries can be found.
'1st parameter: where theme specific resource dictionaries are located
'(used if a resource is not found in the page,
' or application resource dictionaries)

'2nd parameter: where the generic resource dictionary is located
'(used if a resource is not found in the page,
'app, and any theme specific resource dictionaries)
<Assembly: ThemeInfo(ResourceDictionaryLocation.None, ResourceDictionaryLocation.SourceAssembly)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("$guid2$")>

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")>
<Assembly: AssemblyVersion("2017.3.13.1")>
<Assembly: AssemblyFileVersion("2017.3.13.1")>

' I prefer versioning with the year, month, day, build number format like below, but i left the default intact for the template. -Addam
' <Assembly: AssemblyFileVersion("2017.1.10.1")>
