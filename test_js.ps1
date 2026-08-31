$ErrorActionPreference = 'Continue'
$app = New-Object -ComObject Photoshop.Application
Write-Output "Has DoJavaScript: $($app | Get-Member -Name DoJavaScript | Measure-Object | Select-Object -ExpandProperty Count)"
try {
  $result = $app.DoJavaScript("
    var doc = app.documents.add(800, 800, 72, 'JS-Test');
    doc.selection.selectAll();
    var color = new SolidColor();
    color.rgb.red = 255;
    color.rgb.green = 100;
    color.rgb.blue = 50;
    app.foregroundColor = color;
    doc.selection.fill(app.foregroundColor);
    doc.selection.deselect();
    var textLayer = doc.artLayers.add();
    textLayer.kind = LayerKind.TEXT;
    textLayer.textItem.contents = 'Hello from JS';
    textLayer.textItem.size = 48;
    textLayer.textItem.position = [50, 100];
    doc.name + '|' + doc.width.value + 'x' + doc.height.value + '|' + doc.artLayers.length;
  ")
  Write-Output "JS RESULT: $result"
} catch {
  Write-Output "JS FAIL: $($_.Exception.Message)"
}
