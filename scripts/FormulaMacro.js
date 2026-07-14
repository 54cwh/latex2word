function RebuildAll() {
    var count = ActiveDocument.OMaths.Count;
    for (var i = count; i >= 1; i--) {
        try {
            var eq = ActiveDocument.OMaths.Item(i);
            eq.Linearize();
            eq.BuildUp();
        } catch(e) {
            alert("Error on " + i + ": " + e.message);
        }
    }
    alert("Done");
}

function TestLinearize() {
    var eq = ActiveDocument.OMaths.Item(1);
    try {
        eq.Linearize();
        alert("Linearize OK");
    } catch(e) {
        alert("Linearize error: " + e.message);
    }
}
