# Change directory to the required folder 
# Modify values as needed.

foreach ($filepath in (Get-ChildItem -Recurse -Include *index.md).fullname) {

  $fileContent = Get-Content $filePath
  $fileContent[$lineNumber+2] +=  [Environment]::NewLine + 'tags: ["Tech Articles"]'
  $fileContent[$lineNumber+2] += [Environment]::NewLine + 'categories: ["Tech Articles"]'
  $fileContent | Set-Content $filePath 

}



