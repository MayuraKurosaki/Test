Add-Type -AssemblyName System.Windows.Forms

# CSVファイルを読み込むためのヘルパークラス
class CsvHelper {
    [string[]] ReadCsv($filePath) {
        if (-Not (Test-Path $filePath)) {
            throw "File not found: $filePath"
        }
        return Get-Content $filePath
    }
    # CSVファイルの内容をDataTableに変換する
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

# WinFormsのフォームを作成
$form = [System.Windows.Forms.Form]::new()
$form.Text = "CSV Viewer"
$form.Size = [System.Drawing.Size]::new(800, 600)

# DataGridViewを作成
$dataGridView = [System.Windows.Forms.DataGridView]::new()
$dataGridView.Dock = 'Fill'
$dataGridView.AutoSizeColumnsMode = 'Fill'

# ボタンを作成
$button = [System.Windows.Forms.Button]::new()
$button.Text = "Select CSV File"
$button.Dock = 'Top'
$button.Add_Click({
    # CSVファイルを選択するダイアログを表示
    $openFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select a CSV File"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # CSVファイルを読み込み、DataGridViewに表示
        $csvHelper = [CsvHelper]::new()
        $csvData = $csvHelper.ReadCsv($openFileDialog.FileName)
        $dataTable = $csvHelper.ConvertToDataTable($csvData)
        $dataGridView.DataSource = $dataTable
    }
})

# フォームにコントロールを追加
$form.Controls.Add($dataGridView)
$form.Controls.Add($button)

# フォームを表示
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()