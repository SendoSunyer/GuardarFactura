Office.onReady(info => {
    if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
        console.log('Add-in ready');
    }
});

function guardarFactura() {
    // Aquí pots afegir la lògica per guardar la factura.
    console.log('Hello World');
    Office.context.ui.displayDialogAsync('https://localhost:3000/dialog.html', { height: 50, width: 50 });
}
