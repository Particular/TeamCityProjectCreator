Param(
  [Parameter(Mandatory=$true)]
  [string]$ParentID,
  [Parameter(Mandatory=$true)]
  [string]$ProjectName
)

$project_folder = "C:\ProgramData\JetBrains\TeamCity\config\projects"
$user = "tooling"
$template_url = "https://builds.particular.net/httpAuth/app/rest/builds/buildType:Tooling_BuildProcess_StandardNuGet,branch:master,status:SUCCESS/artifacts/content/template.zip"

function Expand-Zip($file, $destination) {
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}

function Replace-Variables($text) {
    $normalized_project_name = Normalize-ProjectName $ProjectName
    $result = $text.Replace("%PARENT_ID%",$ParentID).Replace("%PROJECT_NAME%",$normalized_project_name).Replace("%GUID()%",[Guid]::NewGuid())
    $result
}

function Normalize-ProjectName($project_name) {
    $result = $project_name.Replace(".","_");
    $result
}

$normalized_project_name = Normalize-ProjectName $ProjectName
$full_project_name = "${ParentID}_${normalized_project_name}"
$full_project_path = Join-Path $project_folder $full_project_name
if (Test-Path $full_project_path) {
    throw "Project with specified ID already exists"
}
mkdir $full_project_path

$pass = (Get-Childitem env:TC_PASSWORD).Value
if ($pass -eq $null) {
	throw "TeamCity password for user Tooling need to be specified in environment variable TC_PASSWORD."
}

$pair = "$($user):$($pass)"
$encoded_creds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$basic_auth_value = "Basic $encoded_creds"

$headers = @{
   Authorization = $basic_auth_value
}
Invoke-WebRequest -Uri $template_url -OutFile "template.zip" -Headers $headers

$template_archive = resolve-path ".\template.zip"
if (Test-Path  ".\Template") {
    rmdir ".\Template" -Recurse -Force
}
mkdir "Template"
$destination = Resolve-Path ".\Template"
Expand-Zip -file ($template_archive.Path) -destination ($destination.Path)
Get-ChildItem -Path ".\Template" -Recurse -File | ForEach-Object { 
    $new_name = Replace-Variables $_.FullName
    Rename-Item $_.FullName $new_name
    Replace-Variables (Get-Content $new_name) | Set-Content $new_name
}
Get-ChildItem -Path ".\Template" | ForEach-Object { 
    Copy-Item $_.FullName $full_project_path -Recurse -Force 
}
