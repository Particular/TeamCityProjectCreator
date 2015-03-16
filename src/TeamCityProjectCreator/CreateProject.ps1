Param(
  [string]$ParentID,
  [string]$ProjectName
)

function Expand-Zip($file, $destination) {
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}

$source = "http://builds.particular.net/guestAuth/app/rest/builds/buildType:Tooling_BuildProcess_StandardNuGet,branch:master,status:SUCCESS/artifacts/content/template.zip"
$project_folder = "C:\ProgramData\JetBrains\TeamCity\config\projects"

$full_project_name = "${ParentID}_${ProjectName}"
$full_project_path = Join-Path $project_folder $full_project_name
if (Test-Path $full_project_path) {
    throw "Project with specified ID already exists"
}
mkdir $full_project_path

Invoke-WebRequest $source -OutFile "template.zip"
$template_archive = resolve-path ".\template.zip"
if (Test-Path  ".\Template") {
    rmdir ".\Template" -Recurse -Force
}
mkdir "Template"
$destination = Resolve-Path ".\Template"
Expand-Zip -file ($template_archive.Path) -destination ($destination.Path)
Get-ChildItem -Path ".\Template" -Recurse -File | ForEach-Object { 
    $new_name = $_.FullName.Replace("%PARENT_ID%",$ParentID).Replace("%PROJECT_NAME%",$ProjectName)
    Rename-Item $_.FullName $new_name
    (Get-Content $new_name).Replace("%PARENT_ID%",$ParentID).Replace("%PROJECT_NAME%",$ProjectName) | Set-Content $new_name
}
Get-ChildItem -Path ".\Template" | ForEach-Object { 
    Copy-Item $_.FullName $full_project_path -Recurse -Force 
}
