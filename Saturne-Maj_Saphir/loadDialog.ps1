#[CmdletBinding()]

<#Param(

 [Parameter(Mandatory=$True,Position=1)]

 [string]$XamlPath

)
#>
function Get-ScriptDirectory{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
	$Invocation.PSScriptRoot
}

$Global:xmlWPF = [xml]@"
	<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	Name="Window"  Title="Saturne v4 - Mise à jour Saphir" Height="250" Width="718" ResizeMode="CanMinimize" Icon="$(Get-ScriptDirectory)\Application.ico" WindowStartupLocation = "CenterScreen" >
    <Grid HorizontalAlignment="Left" Width="712">
        <!--En-tête -->
        <Image HorizontalAlignment="Left" Height="82" Margin="10,10,0,0" VerticalAlignment="Top" Width="72" Stretch="None">
            <Image.Source>
                <BitmapImage UriSource="$(Get-ScriptDirectory)\logo-itron.png" />
            </Image.Source>
        </Image>
        <Image Margin="633,10,0,374" Source="$(Get-ScriptDirectory)\icone_saturne.png" RenderTransformOrigin="-0.173,0.521" Stretch="None" Height="48"/>
        <Label Name="Titre" Content="Bienvenue dans l'outil de mise à jour de Saturne." Margin="185,20,203,0" Height="29" Width="332" FontWeight="Bold" RenderTransformOrigin="0.517,1.414" HorizontalAlignment="Center" VerticalAlignment="Top" FontSize="14"/>
        <Label Name="sstitre" Content="                                                    Définissez les options de mise à jour." Margin="63,63,62,0" VerticalAlignment="Top" Height="29" FontWeight="Bold"/>
        <Rectangle Fill="#FFF4F4F5" HorizontalAlignment="Left" Height="64" Margin="10,113,0,0" Stroke="Black" VerticalAlignment="Top" Width="692"/>
        <!--Backup.-->
        <Label Name="Source_path"  Content="Définissez le dossier contenant les nouveaux fichiers :" HorizontalAlignment="Left" Margin="10,113,0,0" VerticalAlignment="Top" Height="27" Width="452" IsEnabled="True"/>
        <TextBox Name ="Source_path_value" HorizontalAlignment="Left" Height="18" Margin="22,138,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="565" IsEnabled="True" Text="C:\Itron\MAJ-Saphir"/>
        <Button Name="Source_path_bouton" Content="Parcourir" HorizontalAlignment="Left" Margin="603,138,0,0" VerticalAlignment="Top" Width="75" Height="18" IsEnabled="True"/>
        <!--Boutons.-->
        <Button Name="Start_bouton" Content="Lancer " HorizontalAlignment="Left" Margin="10,191,0,0" VerticalAlignment="Top" Width="135" Height="20" FontWeight="Bold"/>
        <Button Name="Log_bouton" Content="Ouvrir le fichier de log" Margin="289,191,288,0" VerticalAlignment="Top" Height="20" IsEnabled="False" FontWeight="Bold"/>
        <Button Name="Quit_bouton" Content="Quitter" Margin="567,191,10,0"  Height="20" VerticalAlignment="Top" FontWeight="Bold"/>
    </Grid>
</Window>
"@

 

#Add WPF and Windows Forms assemblies

try{

 Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms

} catch {

 Throw "Failed to load Windows Presentation Framework assemblies."

}

 

#Create the XAML reader using a new XML node reader

$Global:xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))

 

#Create hooks to each named object in the XAML

$xmlWPF.SelectNodes("//*[@Name]") | %{

 Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global

 }