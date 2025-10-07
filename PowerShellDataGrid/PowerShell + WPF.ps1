Add-Type -AssemblyName PresentationFramework

# CSVファイルを読み込むためのヘルパークラス
class CsvHelper {
    [string[]] ReadCsv($filePath) {
        if (-Not (Test-Path $filePath)) {
            throw "File not found: $filePath"
        }
        return Get-Content $filePath
    }

    [System.Data.DataTable] ConvertToDataTable($csvData) {
        $dataTable = [System.Data.DataTable]::new()
        if ($csvData.Length -gt 0) {
            # ヘッダー行を分割してDataTableの列を作成
            $headers = $csvData[0].Split(',')
            foreach ($header in $headers) {
                $dataTable.Columns.Add($header)
            }
            # データ行をDataTableに追加
            for ($i = 1; $i -lt $csvData.Length; $i++) {
                $row = $dataTable.NewRow()
                $values = $csvData[$i].Split(',')
                for ($j = 0; $j -lt $values.Length; $j++) {
                    $row[$j] = $values[$j]
                }
                $dataTable.Rows.Add($row)
            }
        }
        return $dataTable
    }
}

# XAMLコードを定義
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CsvViewer" Width="800" Height="600">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Button Name="selectCsvButton" Grid.Row="0" Content="Select CSV File" Margin="0,0,0,0"/>
        <DataGrid Name="dataGrid" Grid.Row="1" AutoGenerateColumns="True" />
    </Grid>
</Window>
"@

# XAMLを読み込んでウィンドウを作成
$reader = ([System.Xml.XmlNodeReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)
$selectCsvButton = $window.FindName("selectCsvButton")
$dataGrid = $window.FindName("dataGrid")
$selectCsvButton.Add_Click({
    # CSVファイルを選択するダイアログを表示
    $openFileDialog = [Microsoft.Win32.OpenFileDialog]::new()
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select a CSV File"
    if ($openFileDialog.ShowDialog() -eq $true) {
        # CSVファイルを読み込み、DataGridに表示
        $csvHelper = [CsvHelper]::new()
        $csvData = $csvHelper.ReadCsv($openFileDialog.FileName)
        $dataTable = $csvHelper.ConvertToDataTable($csvData)
        $dataGrid.ItemsSource = $dataTable.DefaultView
    }
})

# ウィンドウを表示
$window.ShowDialog() | Out-Null