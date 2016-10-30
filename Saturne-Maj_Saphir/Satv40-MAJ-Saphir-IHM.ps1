##########################################################################
#Name : Satv40-MAJ-Saphir-IHM.ps1                                        #                                
#Description : Installe les elements pour la gestion du Cpt Saphir       #
#Note :                                                                  #
#Author : Julien SIDOT                                                   #
#Contact : Julien.sidot@itron.com                                        #
#ATTENTION : The script must not be modified without the author agreement#
#Version : 1.0                                                           #
##########################################################################
$ver = "1.0"

#Initialisation
        #Defini la taille de la fenetre d'execution pour que les resulats ne soient pas tronqué dans le fichier de log lors de la creation d'un exe ou d'execution par tache planifiée.
            $Shell = $Host.UI.RawUI  
            $size = $Shell.WindowSize 
            $size.width=1000 
            $size.height=60 
            $Shell.WindowSize = $size 
            $size.width=1000 
            $size.height=3000 
            $Shell.BufferSize = $size

        #Define date and log file
            $global:DateStamp = get-date -uformat "%Y-%m-%d@%H-%M-%S"
            $global:RootPath = Get-Location
            If (-not (Test-Path "$RootPath\Logs")) { New-Item -ItemType Directory -Name "Logs" }
            $global:Log = "$RootPath\Logs\Satv40-MAJ-Saphir_$DateStamp.log"

        #Define default value for variables
            # Dossier avec les elements à installer
            $Global:client = "C:\Itron\Satv40-Maj_Saphir".tolower()

#GESTION DE L'IHM

cd $RootPath
	#Required to load the XAML form and create the PowerShell Variables
	.\loadDialog.ps1

	$xamGUI.Add_Closing({$_.Cancel = $true})


#bouton log
    $Log_bouton.add_Click({
    Invoke-Item $log
    })


    #bouton Exit
    $Quit_bouton.add_Click({
    $xamGUI.Dispatcher.InvokeShutdown()
    })

    #bouton Parcourir pour les sources
    $Source_path_bouton.add_Click({
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{SelectedPath = "$Global:client"}
    #[void]$FolderBrowser.ShowDialog()
	if ($FolderBrowser.ShowDialog() -eq "OK")
		{
		$Source_path_value.text = $FolderBrowser.SelectedPath
        $Global:client = $FolderBrowser.SelectedPath
		}
    })


    #Bouton start
    $Start_bouton.add_Click({
    $Start_bouton.IsEnabled = $false
    $Quit_bouton.IsEnabled = $false
    #definitions des variables
    $client = $Source_path_value.text
    .\Satv40-MAJ-Saphir.ps1
    $Quit_bouton.IsEnabled = 'True'
    $Log_bouton.IsEnabled = 'True'
    $Start_bouton.IsEnabled = 'True'
    })
    
    


    
    #Launch the window

    $xamGUI.ShowDialog() | out-null