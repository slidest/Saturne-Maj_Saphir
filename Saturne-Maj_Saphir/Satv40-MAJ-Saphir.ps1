##########################################################################
#Name : Satv40-MAJ-Saphir.ps1                                            #                                
#Description : Installe les elements pour la gestion du Cpt Saphir       #
#Note :                                                                  #
#Author : Julien SIDOT                                                   #
#Contact : Julien.sidot@itron.com                                        #
#ATTENTION : The script must not be modified without the author agreement#
#Version : 1.0                                                           #
##########################################################################
$ver = "1.0"

Function Load-Window {
    
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [void][System.Reflection.Assembly]::LoadWithPartialName('WindowsBase')
    [void][System.Reflection.Assembly]::LoadWithPartialName('system.windows.forms')
    $global:syncHash = [hashtable]::Synchronized(@{})
    $global:newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)          
    $newRunspace.SessionStateProxy.SetVariable("newRunspace",$newRunspace)          
    $global:psCmd = [PowerShell]::Create().AddScript({   
    [xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Name="Window2" Title="" Height="400" Width="725.696" WindowStyle = "none" WindowStartupLocation = "CenterScreen" ResizeMode="NoResize" >
        <Grid>
            <Label Name = "label4"  Content="Saturne v4 - Mise à jour Saphir" HorizontalAlignment="Left" Margin="10,15,0,0" VerticalAlignment="Top" Height="33" Width="258" FontWeight="Bold"/>
            <TextBlock Name = "TextBlock1" HorizontalAlignment="Left" Margin="20,39,0,0" TextWrapping="Wrap" VerticalAlignment="Top" RenderTransformOrigin="0.139,0.627" Width="686"/>
            <ProgressBar Name = "ProgressBar1" Height="20" Margin="20,60,20,0" VerticalAlignment="Top"/>
            <StackPanel Margin="0,105,0,0" >
                <ScrollViewer Name = "scrollviewer" VerticalScrollBarVisibility="Visible"  Height="235" Margin="20,0,20,0">
                    <TextBlock Name = "TextBlock2" TextWrapping = "Wrap" Background="lightgray"/>
                </ScrollViewer >
            </StackPanel>
            <Button Name="Quit_bouton2" Content="Fermer" Margin="557,370,20,0"  Height="20" VerticalAlignment="Top" FontWeight="Bold"/>
        </Grid>
    </Window>
"@
 
        $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )


        $syncHash.IsClosed = $False
        $syncHash.Error = $Error

        #$reader=(New-Object System.Xml.XmlNodeReader $xaml)
        #$syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
        $syncHash.ProgressBar1 = $syncHash.Window.FindName("ProgressBar1")
        $syncHash.TextBlock1 = $syncHash.Window.FindName("TextBlock1")
        $syncHash.TextBlock2 = $syncHash.Window.FindName("TextBlock2")
        $syncHash.scrollviewer = $syncHash.Window.FindName("scrollviewer")
        $syncHash.button = $syncHash.Window.FindName("Quit_bouton2")
        $syncHash.button.IsEnabled = $false
        $syncHash.Window.Add_Closing( { If ($syncHash.IsClosed -ne $True) { $_.Cancel = $True; $syncHash.IsClosed = $True} } )
        $syncHash.button.Add_Click( { $Global:syncHash.Window.Dispatcher.invoke(“Normal”,[action]{ $Global:syncHash.Window.Close() })} )
        #$syncHash.button.Add_Click({[System.Windows.Forms.MessageBox]::Show("Hello World." , "My Dialog Box") })
        $syncHash.Window.ShowDialog() | Out-Null
    })
    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
    While (!($syncHash.Window.IsInitialized)) { Start-Sleep -Milliseconds 1 }

} # End Function Load-Window

Function Update-Window {
    Param (
        $Control,
        $Property,
        $Value,
        [switch]$AppendContent
    )

    $syncHash.IsRunning = $true


    #Write-Warning "$Control . $Property = $Value" 
    If ($newRunspace.RunspaceStateInfo.State -ne 'Opened') { Return $False }

    # User is closing the window.
    If ($syncHash.IsClosed) { $Property = "Close" }

    # This is kind of a hack, there may be a better way to do this
    If ($Property -eq "Close") {

        $syncHash.IsClosed = $True
        If($PSVersionTable.PSVersion.Major -ge 2) 
            {
                $Global:syncHash.$Control.Dispatcher.invoke(“Normal”,[action]{ $Global:syncHash.Window.Close() })
            } 
        Else 
            {
                $Global:syncHash.Window.Dispatcher.invoke([action]{ $Global:syncHash.Window.Close() },"Normal")
            }

        $newRunspace.Close()
        Return $False
    }
  
    # This updates the control based on the parameters passed to the function
    If (($Control) -and ($Property) -and ($Value)) {
        
        if ($control -eq "TextBlock2")
            {
                $syncHash.$Control.Dispatcher.Invoke("Normal",[action]{                       
                            $Run = New-Object System.Windows.Documents.Run
                            if ($value.Contains("Succes")) {$Run.Foreground = "Green"}
                            elseif ($value.Contains("Info")) {$Run.Foreground = "Yellow"}
                            elseif ($value.Contains("Echec")) {$Run.Foreground = "Red"}
                            elseif ($value.Contains("!!! ATTENTION !!!")) {$Run.Foreground = "Yellow"}
                            else {$Run.Foreground = "Black"}
                            $Run.Text = ("{0}" -f $value)
                            $syncHash.$Control.Inlines.Add($Run)
                            $syncHash.$Control.Inlines.Add((New-Object System.Windows.Documents.LineBreak))                                                  
                        })
                $syncHash.scrollviewer.Dispatcher.Invoke("Normal",[action]{
                            $syncHash.scrollviewer.ScrollToEnd()
                        })
            }

        else
            {
                If($PSVersionTable.PSVersion.Major -ge 2) 
                    {
                        $Global:syncHash.$Control.Dispatcher.invoke(“Normal”,[action]{ 
                        If ($PSBoundParameters['AppendContent']) { $syncHash.$Control.AddText($Value) } 
                                                            Else { $syncHash.$Control.$Property = $Value }
                        })
                    } 
                Else 
                    {
                        $Global:syncHash.Window.Dispatcher.invoke([action]{ 
                        If ($PSBoundParameters['AppendContent']) { $syncHash.$Control.Addtext($Value) } 
                        Else { $syncHash.$Control.$Property = $Value }
                        },"Normal")
                    }
            }

    }

    Return $True

} # End Function Update-Window

Function Close-Window {

    Update-Window Window Close | Out-Null
} # End Function Close-Window


Function Update-ProgressBar {
    Param (
        [int]$Percent1,
        [string]$TextBlock1,
        [string]$TextBlock2,
        [string]$button
    )

    If ($TextBlock1) { $Result = Update-Window TextBlock1 Text $TextBlock1 }
    If ($TextBlock2) { $Result = Update-Window TextBlock2 Text $TextBlock2 }
    If ($Percent1)   { $Result = Update-Window ProgressBar1 Value $Percent1 }
    Return $Result

} #End Function Update-ProgressBar




# Below : the exectution script and 
# load new window
Load-Window




$Result = Update-Window


Write-Output "************************************************************************************************************************************" >> $Log
Write-Output "                                         Mise à jour de Saturne v40 pour le compteur Saphir"                                          >> $Log
Write-Output "************************************************************************************************************************************" >> $Log

#Nombre d'action : init, ident comp, MAJ File server, MAJ WebSATURNE, MAJ Svc Comm
$act = 5
$count = 0

$Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Initialisation"

    #Initialisation
        #Define date
            $StartDate = get-date -uformat "%d/%m/%Y at %H:%M:%S"

	    #Load PS Modules
            if ((Get-WmiObject Win32_OperatingSystem).Name.Contains('Server 2008 R2'))
            {
                Import-Module servermanager
            }
        #Load Web components (clients and services)
            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration")
            $iisVersion = Get-ItemProperty "HKLM:\software\microsoft\InetStp";
            if ($iisVersion.MajorVersion -ge 8)
                {
                    Import-Module WebAdministration;
                } 
            elseif ($iisVersion.MajorVersion -eq 7)
            {
                if ($iisVersion.MinorVersion -ge 5)
                {
                    Import-Module WebAdministration;
                }           
                else
                {
                    if (-not (Get-PSSnapIn | Where {$_.Name -eq "WebAdministration";})) {
                        Add-PSSnapIn WebAdministration;
                    }
                }
            }

        # controle si le serveur est 64 ou 32 bits pour les clé de registres
            if ([intptr]::Size -eq 8)
                {
                    $pre_key = "HKLM:\SOFTWARE\Wow6432Node"
                }
            elseif ([intptr]::Size -eq 4)
                {
                    $pre_key = "HKLM:\SOFTWARE"
                }

        #Définition des fonctions
            # Pour copier un fichier ou dossier
                function copie ($src, $dest, $Log)
                    {
                        Copy-Item "$src" "$dest"  -force -Recurse
                        if ($? -eq "True")
                            {
                                $result = Update-ProgressBar -TextBlock2 "Succes de la copie de `"$src`" dans `"$dest`"."
                                Write-Output "Succes de la copie de `"$src`" dans `"$dest`"." >>$Log
                            }
                        else 
                            {
                                $result = Update-ProgressBar -TextBlock2 "Echec de la copie de `"$src`" dans `"$dest`"."
                                Write-Output "Echec de la copie de `"$src`" dans `"$dest`"." >>$Log
                            }
                    }


            #Pour arreter un service
                function SvcStop ($Svc, $log)
                    {
                        Stop-Service $svc -PassThru
                        if ($? -eq "True")
                            {
                                $result = Update-ProgressBar -TextBlock2 "Succes de l'arret du service `"$svc`"."
                                Write-Output "Succes de l'arret du service `"$svc`"." >>$Log
                            }
                        else 
                            {
                                $result = Update-ProgressBar -TextBlock2 "Echec de l'arret du service `"$svc`"."
                                Write-Output "Echec de l'arret du service `"$svc`"." >>$Log
                            }
                    }

            #Pour démarrer un service
                function SvcStart ($Svc, $log)
                    {
                        Start-Service $svc -PassThru
                        if ($? -eq "True")
                            {
                                $result = Update-ProgressBar -TextBlock2 "Succes du démarrage du service `"$svc`"."
                                Write-Output "Succes du démarrage du service `"$svc`"." >>$Log
                            }
                        else 
                            {
                                $result = Update-ProgressBar -TextBlock2 "Echec du démarrage du service `"$svc`"."
                                Write-Output "Echec du démarrage du service `"$svc`"." >>$Log
                            }
                    }
                
            
            #Pour creer des fichiers ou dossiers ou des clés de registre
                function create_Item($fld, $type, $log)
                    {
                        if ((test-path "$fld") -eq $false)
                            {
                                new-item "$fld" -type $type
                                if ($? -eq "True")
			                        {
				                        $result = Update-ProgressBar -TextBlock2 "Succes de la création de `"$fld`"."
				                        Write-output "Succes de la création de `"$fld`"." >> $Log
			                        }
		                        else
			                        {
                                        $result = Update-ProgressBar -TextBlock2 "Echec de la création de `"$fld`"." 
				                        Write-output "Echec de la création de `"$fld`"." >> $Log
			                        }
                            }
                        else
                            {
                                $result = Update-ProgressBar -TextBlock2 "Info : l'objet `"$fld`" existe déja." 
				                Write-output "Info : l'objet `"$fld`" existe déja." >> $Log
                            }
                    }

    #Traitement
        $count++
        $Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Identification des composants Saturne installés"
        #Identification des composants Saturne installés
            $tmp = Get-ItemProperty "$pre_key\Microsoft\Windows\CurrentVersion\Uninstall\*" | where {$_.DisplayName -ne $null}
       
            $InstalledComp=@()
            foreach ($i in $tmp){
                if (($i.DisplayName.contains('Saturne')) -or ($i.DisplayName.contains('Asaïs - Serveur de fichiers')) -or ($i.DisplayName.contains('File Server')) -or ($i.DisplayName.contains('MS21Server')))
                    {
                    if (!($i.DisplayName.contains('Saturne Mobile')))
                        {
                            $InstalledComp += $i
                        }
                    }
                }

        #Si le serveur de fichier est installé
            $count++
            $Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Mise à jour du serveur de fichier"
            Write-Output "************************************************************************************************************************************" >> $Log
            Write-Output "                                         Mise à jour du serveur de fichier                 "                                          >> $Log
            Write-Output "************************************************************************************************************************************" >> $Log
            if (($InstalledComp.Displayname -match 'Asaïs - Serveur de fichiers') -or ($InstalledComp.Displayname -match 'File Server'))
                {

                    #arret du service
                    SvcStop 'Asais File Server' $log

                    #Identifier le chemin vers le dossier Root
                    $Rpath = (Get-ItemProperty "$pre_key\Asais\FileServer").RootDir

                    #Renomage de l'ancien fichier de licence
                        $lic = (Get-ChildItem "$Rpath\Common\Licences\*" -Include "*.lic")
                        $namelic= ($lic.Name).substring(0, ($lic.Name).Length -4) + "_$DateStamp.old"
                        Rename-Item $lic.FullName $namelic

                    #copie de la nouvelle licence
                        $dir = $lic.DirectoryName + "\"
                        $newlic = (Get-ChildItem "$client\*" -Include "*.lic").FullName
                        copie $newlic $dir $Log
                    #Copie du nouveaux TypesEqp et du SvcDef
                        $src = "$client\Data\Root\*"
                        $dir = "$Rpath\"
                        copie $src $dir $log

                    #Démarrage du service
                        SvcStart 'Asais File Server' $log
                }
            else
                {
                    $result = Update-ProgressBar -TextBlock2 "Infos :  le composant `"Asaïs - Serveur de fichiers`" n'est pas présent sur le serveur" 
                }


        #Si WebSaturne installé
            $count++
            $Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Mise à jour du WebSaturne"
            Write-Output "************************************************************************************************************************************" >> $Log
            Write-Output "                                         Mise à jour du WebSaturne                        "                                          >> $Log
            Write-Output "************************************************************************************************************************************" >> $Log
            if ($InstalledComp.Displayname -match 'Saturne - Intranet')
                {
                    #identification du dossier d'installation de WebSaturne
                        $com = (((New-Object Microsoft.Web.Administration.ServerManager).Sites).Applications) | where { $_.Path -eq "/WebSaturne"}
                        $com2 = $com.VirtualDirectories
                        $com3 = $com2.PhysicalPath
                    
                    #Renomage de l'ancien fichier declinaisons.xml
                        $dec = (Get-Item "$com3\XSLINDEX\Declinaisons.xml")
                        $namedec= ($dec.Name).substring(0, ($dec.Name).Length -4) + "_$DateStamp.old"
                        Rename-Item $dec.FullName $namedec

                    #Copie du nouveau declinaisons.xml
                        $dir = $dec.DirectoryName + "\"
                        $newdec = (Get-ChildItem "$client\Data\WEB\WebSaturne\XSLINDEX\Declinaisons.xml").FullName
                        copie $newdec $dir $Log
                    #IISreset
                        & {iisreset >>$log}
                }
            else
                {
                    $result = Update-ProgressBar -TextBlock2 "Infos :  le composant `"Saturne - Intranet`" n'est pas présent sur le serveur" 
                }

        #Si Srv com installé
            $count++
            $Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Mise à jour du Serveur de communication"
            Write-Output "************************************************************************************************************************************" >> $Log
            Write-Output "                                         Mise à jour du Communication server               "                                          >> $Log
            Write-Output "************************************************************************************************************************************" >> $Log
            if ($InstalledComp.Displayname -match 'Saturne Communication server Service')
                {
                    $svc = 'SatPostesvcs'
                    #arret du service
                        SvcStop $svc $log
                    #Determination du dossier d'install du service
                        $Service = gwmi win32_service| where {$_.Name -eq "$svc"}
                        $path = $Service.Pathname
                        $pos = $path.LastIndexOf("\")
                        $path = $path.Substring(1,$pos-1)
                    #backup des anciens protos
                        $dest = "$path\Backup_$DateStamp"
                        create_Item "$path\Backup_$DateStamp\fr" 'Directory' $log
                        $list = @()
                        (Get-ChildItem "$client\Data\Drivers" | where-object { -not $_.PSIsContainer}).name | where { -not ($_ -match "esources.dll") } | foreach {
                            if (get-item "$path\$_")
                                {
                                    $file = (get-item "$path\$_").fullname
                                    Move-Item $file "$path\Backup_$DateStamp\"
                                }
                            }
                        (Get-ChildItem "$client\Data\Drivers\fr" | where-object { -not $_.PSIsContainer}).name  | foreach {
                            if (get-item "$path\fr\$_")
                                {
                                    $file = (get-item "$path\fr\$_").fullname
                                    Move-Item $file "$path\Backup_$DateStamp\fr\"
                                }
                            }
                    #copie des nouveaux protos
                        copie "$client\Data\Drivers\*" "$path\" $log
                    #Démarrage du service
                        SvcStart 'SatPostesvcs' $log
                }
            else
                {
                    $result = Update-ProgressBar -TextBlock2 "Infos :  le composant `"Saturne Communication server Service`" n'est pas présent sur le serveur" 
                }

 
        #activation du bouton quitter pour la fermeture de la fenetre
            $count++
            $Result = Update-ProgressBar -Percent1 (($count/$act)*100) -TextBlock1 "Terminé" 
            $result = Update-ProgressBar -TextBlock2 "************************************************************************************************************************************"
            $result = Update-ProgressBar -TextBlock2 "Traitement de Saturne terminé. Controlez le fichier de trace."
            Write-Output "************************************************************************************************************************************" >> $Log
            Write-Output "************************************************************************************************************************************" >> $Log
            Write-Output "                                         Mise à jour terminée                              "                                          >> $Log
            
            $Result = Update-Window button IsEnabled "True" 