$folders = Get-ChildItem -Path "C:\users\rjefferson\Documents\test" -Directory -Force -ErrorAction SilentlyContinue


$addUserToFolder = try{
                        foreach ($filePath in $folders){
                                $path = $filePath | Select-Object FullName -ExpandProperty FullName #Replace with whatever file you want to do this to.
                                $findUser = $filePath | Select-Object Name -ExpandProperty Name
                                $user = "ptown.local\$findUser" #User account to grant permisions too.
                                #$addAdmin = "ptown.local\Domain Admin"
                                $Rights = "FullControl, ReadAndExecute, ListDirectory" #Comma seperated list.
                                $InheritSettings = "Containerinherit, ObjectInherit" #Controls how permissions are inherited by children
                                $PropogationSettings = "None" #Usually set to none but can setup rules that only apply to children.
                                $RuleType = "Allow" #Allow or Deny.

                                $acl = Get-Acl $path
                                $perm = $user, $Rights, $InheritSettings, $PropogationSettings, $RuleType
                                $rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $perm
                                $acl.SetAccessRule($rule)
                                $acl | Set-Acl -Path $path
                                }
} catch { 
        #cathes the user not within AD and writes the folder name in the text file.
        Add-Content C:\users\rjefferson\Documents\test\HDriveCleanup.txt "$findUser Folder needs to be Deleted."
}

$addDomainAdmin = foreach ($filePaths in $folders){
                    $path = $filePaths | Select-Object FullName -ExpandProperty FullName #Replace with whatever file you want to do this to
                    $user = "ptown.local\Domain Admins" #User account to grant permisions too.
                    #$addAdmin = "ptown.local\Domain Admin"
                    $Rights = "FullControl, ReadAndExecute, ListDirectory" #Comma seperated list.
                    $InheritSettings = "Containerinherit, ObjectInherit" #Controls how permissions are inherited by children
                    $PropogationSettings = "None" #Usually set to none but can setup rules that only apply to children.
                    $RuleType = "Allow" #Allow or Deny.

                    $acl = Get-Acl $path
                    $perm = $user, $Rights, $InheritSettings, $PropogationSettings, $RuleType
                    $rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $perm
                    $acl.SetAccessRule($rule)
                    $acl | Set-Acl -Path $path
}

$folder = Get-ChildItem -Path "C:\users\rjefferson\Documents\test" -Recurse -Directory -Force -ErrorAction SilentlyContinue
$removeUser= foreach ($filePath in $folder){
                    $path = $filePath | Select-Object FullName -ExpandProperty FullName
                    cacls $path /e /p Everyone:n
}

$addUserToFolder
$addDomainAdmin
$removeUser