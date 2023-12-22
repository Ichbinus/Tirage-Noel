try {
    $assembly_list = "PresentationFramework","System.Windows.Forms","PresentationCore","WindowsBase","System.Xaml","UIAutomationClient","UIAutomationTypes","WindowsFormsIntegration","System","System.Core","mscorlib","System.Management.Automation","System.Threading"
    foreach ($assembly in $assembly_list)
    {
        # Chargement des assemblies
        Add-Type -AssemblyName $assembly
        # Test des assemblies chargées
        $loadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()| Where-Object { $_.GetName().Name -eq $assembly } 
        $assemblyLoaded = $loadedAssemblies.FullName -contains $assembly
        if ($null -ne $assemblyLoaded) {
            $logassembly = "L'assembly '$assembly' est chargé correctement."
            ###WriteToLogFile0 $logassembly
        } else {
            $logassembly = "L'assembly '$assembly' n'est pas chargé correctement."
            ###WriteToLogFile0 $logassembly
        }
    }
    $loaded_assembly_list = [System.AppDomain]::CurrentDomain.GetAssemblies()| Select-Object -Property FullName 
    foreach ($loaded_assembly in $loaded_assembly_list)
    {
        $logloadedassembly = "L'assembly suivante est chargé: '$loaded_assembly'"
        ###WriteToLogFile0 $logloadedassembly
    }
}
catch {
    $logwpf = "ligne 55 - une exeption s'est produite $_.Exception.Message "
    ###WriteToLogFile0 $logwpf
}
[xml]$Fenetre_principale = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        Title="MainWindow" Height="300" Width="485">
    <Grid Background="#FFC1C3CB">
        <Button Name="Start" Content="Tirage au sort" HorizontalAlignment="Left" FontSize="20" FontWeight="Bold" FontFamily="Comic Sans MS" Margin="34,10,0,0" VerticalAlignment="Top" Width="174" Height="46" />
        <Button Name="Accepter" Content="Tirage ok" HorizontalAlignment="Left" FontSize="20" FontWeight="Bold" FontFamily="Comic Sans MS" Margin="238,10,0,0" VerticalAlignment="Top" Width="174" Height="46" />
        <GroupBox Header="Qui offre à qui?" FontSize="15" FontWeight="Bold" FontFamily="Comic Sans MS" BorderBrush="CornFlowerBlue" Margin="10,61,10,10">
            <DataGrid SelectionMode="Single" Name="Grille_tirage" ItemsSource="{Binding}" Height="323">
                <DataGrid.Columns >
                    <DataGridTextColumn Header="Participant" Binding="{Binding Participant}"/>
                    <DataGridTextColumn Header="L'Année Dernière" Binding="{Binding Année_prec}"/>
                    <DataGridTextColumn Header="Cette Année" Binding="{Binding Année_N}"/>
                    <DataGridTextColumn Header="Tirage Valide" Binding="{Binding Valide}"/>
                </DataGrid.Columns>
                <DataGrid.RowStyle>
                    <Style TargetType="DataGridRow">
                        <Setter Property="IsHitTestVisible" Value="True"/>
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding Valide}" Value="Pas Ok">
                                <Setter Property="Background" Value="#E14D57"></Setter>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding Valide}" Value="Ok">
                                <Setter Property="Background" Value="#71B37C"></Setter>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </DataGrid.RowStyle>
            </DataGrid>
        </GroupBox>
    </Grid>
</Window>
"@

$FormFenetre_principale = (New-Object System.Xml.XmlNodeReader $Fenetre_principale)
$Window = [Windows.Markup.XamlReader]::Load($FormFenetre_principale)

## déclaration des contrôles
$Datagrid_tirage = $Window.FindName("Grille_tirage")

$BTN_Start = $Window.FindName("Start")
$BTN_Accepter = $Window.FindName("Accepter")
$BTN_Accepter.IsEnabled=$false

## déclaration des variables
$Historique_tirage = ".\historique_tirage.xml"
$currentYear  =(Get-Date).Year
$currentyearnode = "Année$($currentYear)"
$previousYear = "Année$($currentYear - 1)"

## déclaration des fonctions
function MàJ_grille_tirage {
    $tirage_Valeur = New-Object PSObject
    $tirage_Valeur = $tirage_Valeur | Add-Member NoteProperty Participant $tireur.Nom -passthru	
    $tirage_Valeur = $tirage_Valeur | Add-Member NoteProperty Année_prec $tireur.$previousYear -passthru
    $tirage_Valeur = $tirage_Valeur | Add-Member NoteProperty Année_N $tirage -passthru
    #validation du tirage pour coloration lignes
    if (($tirage -eq $tireur.Nom) -or ($tirage -eq $tireur.$previousYear)) {
        $validation_tirage = "Pas Ok"
    }else {
        $validation_tirage = "Ok"
    }
    $tirage_Valeur = $tirage_Valeur | Add-Member NoteProperty Valide $validation_tirage -passthru
    $Window.Dispatcher.invoke([action]{
    $Datagrid_tirage.Items.Add($tirage_Valeur) > $null 
    })
}

function MàJ_tirage_xml {
    foreach ($item in $Datagrid_tirage.Items) {
        $participant = $item.Participant
        Write-Host $participant  
        $anneeN = $item.Année_N
        Write-Host $anneeN 
        $historiqueXML = [xml](Get-Content $Historique_tirage)
        # Créer un nouvel élément avec le tirage de l'année pour le participant
        $newtirage=$historiqueXML.Tirage.$participant.AppendChild($historiqueXML.CreateElement($currentyearnode)) 
        $newtirage_Value = $newtirage.AppendChild($historiqueXML.CreateTextNode($anneeN));
        # Ajouter le résultat du tirage au participant
        $historiqueXML.Save($Historique_tirage)
    }
}

function tirage_au_sort {
    $historiqueXML = [xml](get-content $Historique_tirage)
    #$Liste_participants = $historiqueXML.SelectNodes("//Participant/joueur").InnerText
    $Liste_participants = $historiqueXML.Tirage.Participant
    #$Liste_participants
    $Liste_nom_participants =@()
    foreach($participant in $Liste_participants)
    {
        $Liste_nom_participants += $participant.Nom
    }
    $ParticipantsRestants = $Liste_nom_participants
    $ParticipantsRestants
    $Liste_participants
    foreach($tireur in $Liste_participants)
    {
        #Write-Host "tirage de: $participant"
        $tirage =""
        $tirage = Get-Random -InputObject $ParticipantsRestants
        #Write-Host "$offrant devra offrir un cadeau à $tirage"
        $ParticipantsRestants = $ParticipantsRestants -ne $tirage
        MàJ_grille_tirage 
    }
}

################################################################################################################################################
################################################################################################################################################
########################################################## Lancement tirage au sort ############################################################ 
################################################################################################################################################
################################################################################################################################################
$BTN_Start.add_click({
    $BTN_Accepter.IsEnabled=$true
    $Datagrid_tirage.Items.Clear()
    $Datagrid_tirage.Items.Refresh()
    tirage_au_sort
})


################################################################################################################################################
################################################################################################################################################
########################################################## Valider tirage au sort ############################################################ 
################################################################################################################################################
################################################################################################################################################
$BTN_Accepter.add_click({
    $BTN_Accepter.IsEnabled=$false
    $Datagrid_tirage.Items.Clear()
    $Datagrid_tirage.Items.Refresh()
    MàJ_tirage_xml
})

$Window.ShowDialog() | Out-Null

Read-Host -Prompt "Press Enter to exit"