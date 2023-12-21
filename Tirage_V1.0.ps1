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
        Title="MainWindow" Height="450" Width="485">
    <Grid Background="#FFC1C3CB">
        <Button Name="Start" Content="Tirage au sort" HorizontalAlignment="Center" FontSize="20" FontWeight="Bold" FontFamily="Comic Sans MS" Margin="0,10,0,0" VerticalAlignment="Top" Width="174" Height="46" />
        <GroupBox Header="Qui offre à qui?" FontSize="16" FontWeight="Bold" FontFamily="Comic Sans MS" BorderBrush="CornFlowerBlue" Margin="7,61,7,93">
            <DataGrid Name="ConnexionGrid" ItemsSource="{Binding}" Height="323">
                <DataGrid.Columns >
                    <DataGridTextColumn Header="Participant" Binding="{Binding Participant}"/>
                    <DataGridTextColumn Header="Cette Année" Binding="{Binding Année_N}"/>
                    <DataGridTextColumn Header="L'Année Dernière" Binding="{Binding Année_N-1}"/>
                </DataGrid.Columns>
            </DataGrid>
        </GroupBox>
        <Button Name="Accepter" Content="Tirage ok" HorizontalAlignment="Left" FontSize="20" FontWeight="Bold" FontFamily="Comic Sans MS" Margin="34,358,0,0" VerticalAlignment="Top" Width="174" Height="46" />
        <Button Name="Relancer" Content="Relancer tirage" HorizontalAlignment="Left" FontSize="20" FontWeight="Bold" FontFamily="Comic Sans MS" Margin="273,358,0,0" VerticalAlignment="Top" Width="174" Height="46" />
    </Grid>
</Window>
"@

$FormFenetre_principale = (New-Object System.Xml.XmlNodeReader $Fenetre_principale)
$Window = [Windows.Markup.XamlReader]::Load($FormFenetre_principale)

$Window.ShowDialog() | Out-Null
