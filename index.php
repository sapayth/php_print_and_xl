<?php

ob_start();	// start output buffer
?>
<div id="rptTableArea">
	<div class="heading" style="width:60%; margin: 20px auto; text-align: center;">
		<h2>Some sample report</h2>
	</div>
	<div class="report-area" style="width:60%; margin: 20px auto; text-align: center;">
		<table width="100%" cellspacing="0" border="1">
			<thead>
				<th>SL</th>
				<th>Merchandiser</th>
				<th>Job</th>
				<th>PO</th>
				<th>Style Ref.</th>
				<th>Shipment</th>
			</thead>
			<tbody>
				<tr>
					<td>1</td>
					<td>Md. Shamsul Arefin</td>
					<td>276</td>
					<td>po-276-04-20-a4</td>
					<td>gangnam style</td>
					<td>18-MAY-20</td>
				</tr>
				<tr>
					<td>1</td>
					<td>Abdul Baten</td>
					<td>133</td>
					<td>po-03</td>
					<td>Min ship</td>
					<td>07-MAY-20</td>
				</tr>
			</tbody>
		</table>
	</div>
</div>
<div style="width:60%; margin: 20px auto; text-align: center;">

	<a href="#" id="btnExcel" download><input type="button" value="Excel Preview" name="excel" class="formbutton" style="width:100px;"/></a>
	<a href="#" id="btnPrint" onclick="printPreview();"><input type="button" value="HTML Preview" name="print" class="formbutton" style="width:100px"/></a>
</div>

<?php

// first delete all the previous .xls
foreach (glob("*.xls") as $filename) {       
    @unlink($filename);
}

$fileName='report_'.time().'.xls';
$create_new_excel = fopen($fileName, 'w');
$is_created = fwrite($create_new_excel, ob_get_contents());
?>

<script>
	window.onload = function() {
		document.getElementById('btnExcel').href = "<?php echo $fileName; ?>";
	}
	
</script>

<?php
ob_end_flush();	// end and flush output buffer
?>

<script>
	function printPreview() {
		var w = window.open('', '_blank');
        var d = w.document.open();
        d.write ('<!DOCTYPE HTML>'+'<html><head><title>Print Preview</title></head><body>'+document.getElementById('rptTableArea').innerHTML+'</body</html>');
        d.close();
	}
</script>