var X = XLS;
var XW = {
	/* worker message */
	msg: 'xls',
	/* worker scripts */
	rABS: './scripts/xlsxworker2.js',
	norABS: './scripts/xlsxworker1.js',
	noxfer: './scripts/xlsxworker.js'
};

var rABS = false;//typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
var use_worker = false;//typeof Worker !== 'undefined';
var transferable = use_worker;

var wtf_mode = false;

var respuestasAlumnos = { PreA:[], PreB:[], PostA:[], PostB:[] };

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
		o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

function ab2str(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint16Array(data.slice(l*w,l*w+w)));
		o+=String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w)));
	return o;
}

function s2ab(s) {
	var b = new ArrayBuffer(s.length*2), v = new Uint16Array(b);
	for (var i=0; i != s.length; ++i) v[i] = s.charCodeAt(i);
		return [v, b];
}

function xw_noxfer(data, cb) {
	var worker = new Worker(XW.noxfer);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			case XW.msg: cb(JSON.parse(e.data.d)); break;
		}
	};
	var arr = rABS ? data : btoa(fixdata(data));
	worker.postMessage({d:arr,b:rABS});
}

function xw_xfer(data, cb) {
	var worker = new Worker(rABS ? XW.rABS : XW.norABS);
	worker.onmessage = function(e) {
		switch(e.data.t) {
			case 'ready': break;
			case 'e': console.error(e.data.d); break;
			default: xx=ab2str(e.data).replace(/\n/g,"\\n").replace(/\r/g,"\\r"); console.log("done"); cb(JSON.parse(xx)); break;
		}
	};
	if(rABS) {
		var val = s2ab(data);
		worker.postMessage(val[1], [val[1]]);
	} else {
		worker.postMessage(data, [data]);
	}
}

function xw(data, cb) {
	if(transferable) xw_xfer(data, cb);
	else xw_noxfer(data, cb);
}


function leerTexto(hoja,celda){
	if (!hoja[celda]) return "";
	return hoja[celda].v;
}

function enum2From(x,y){
	var ns = [];
	for(var i=x; i<=y; i+=2){
		ns.push(i);
	}
	return ns;
}

function leerChoice(hoja,n1,n2){
	var respuestas = enum2From(n1,n2)
		.filter(function(n){ 
			return leerTexto(hoja,'C'+n) !== ""})
		.map(function(n){
			return leerTexto(hoja,'D'+n) + ","});
	if(respuestas[respuestas.length-1]==" Otro: ," || leerTexto(hoja,'E150') != ""){
		respuestas.pop();
		respuestas.push(" Otro: " + leerTexto(hoja,'E'+n2))
	};

	var concatenacion = "";
	for (var i=0; i < respuestas.length; i++){
		concatenacion += respuestas[i];
	};

	if(concatenacion.slice(-1)==","){
		concatenacion = concatenacion.slice(0,concatenacion.length-1);
	}
	return concatenacion;
}

function idTest(workbook){
	var id = leerTexto(workbook.Sheets[workbook.SheetNames[0]], 'A22')
	if(id ==="") return leerTexto(workbook.Sheets[workbook.SheetNames[0]], 'A21'); //El puto ods lee mal
	return id;
}

function listaPara(workbook){ 
	var hoja1 = workbook.Sheets[workbook.SheetNames[0]];
	var listas = {};
	listas.PreA = function(){return[ 
		leerTexto(hoja1,'B4'),
        leerTexto(hoja1,'B8'),
        leerTexto(hoja1,'B12'),
        leerTexto(hoja1,'B15'),
        leerTexto(hoja1,'B18'),
        leerTexto(hoja1,'B41'),
        leerTexto(hoja1,'B52'), //Pregunta 1
        leerTexto(hoja1,'B76'),
        leerTexto(hoja1,'B109'),
        leerChoice(hoja1,126,150), // Pregunta 4
        leerTexto(hoja1,'E155'),
        leerTexto(hoja1,'B169'),
        leerTexto(hoja1,'B174')]};
	listas.PreB = function(){return[
		leerTexto(hoja1,'B4'),
        leerTexto(hoja1,'B8'),
        leerTexto(hoja1,'B12'),
        leerTexto(hoja1,'B15'),
        leerTexto(hoja1,'B18'),
        leerTexto(hoja1,'B41'),
        leerTexto(hoja1,'B52'), //Pregunta 1
        leerTexto(hoja1,'B76'),
        leerTexto(hoja1,'B109'),
        leerChoice(hoja1,126,150), // Pregunta 4
        leerTexto(hoja1,'E155'),
        leerTexto(hoja1,'B169'),
        leerTexto(hoja1,'B176')]};
	listas.PostA = function(){return[ 
		leerTexto(hoja1,'B4'),
        leerTexto(hoja1,'B8'),
        leerTexto(hoja1,'B12'),
        leerTexto(hoja1,'B15'),
        leerTexto(hoja1,'B18'),
        leerTexto(hoja1,'B41'),
        leerTexto(hoja1,'B52'), //Pregunta 1
        leerTexto(hoja1,'B76'),
        leerTexto(hoja1,'B109'),
        leerChoice(hoja1,126,150), // Pregunta 4
        leerTexto(hoja1,'E155'),
        leerTexto(hoja1,'B169'),
        leerTexto(hoja1,'B193'),
		leerTexto(hoja1,'B209'),
		leerTexto(hoja1,'E235'),
		leerTexto(hoja1,'B188')]};
	listas.PostB = function(){return[ 
		leerTexto(hoja1,'B4'),
        leerTexto(hoja1,'B8'),
        leerTexto(hoja1,'B12'),
        leerTexto(hoja1,'B15'),
        leerTexto(hoja1,'B18'),
        leerTexto(hoja1,'B41'),
        leerTexto(hoja1,'B52'), //Pregunta 1
        leerTexto(hoja1,'B76'),
        leerTexto(hoja1,'B109'),
        leerChoice(hoja1,126,150), // Pregunta 4
        leerTexto(hoja1,'E155'),
        leerTexto(hoja1,'B169'),
        leerTexto(hoja1,'B193'),
		leerTexto(hoja1,'B188'),
		leerTexto(hoja1,'B209'),
		leerTexto(hoja1,'E235')]};

	return listas[idTest(workbook)]();
}

function crearHeader(tabla,workbook){
	document.getElementsByClassName(idTest(workbook))[0].hidden=false;

	var hoja2 = workbook.Sheets[workbook.SheetNames[1]];
	var columnas = [];
	var celdas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	/*for(var i=0; celdas.length > i && (leerTexto(hoja2,celdas[i]+1) !== ""); i++){
		columnas.push(leerTexto(hoja2,celdas[i]+1));
	};
	agregarFila(tabla.createTHead(),columnas);*/
	tabla.appendChild(document.createElement('tbody'));
	respuestasAlumnos[idTest(workbook)].push(columnas);
}

function tablaPara(workbook){
	var table = document.getElementById(idTest(workbook));
	if(!table.tHead) crearHeader(table,workbook);
	return table;
}

function agregarFila(table,lista){
	var row = table.insertRow(table.rows.length);
	for(var i=0; i< lista.length; i++){
		row.insertCell(i).innerHTML = lista[i];
	};
}

function process_wb(wb) {
	if(use_worker) XLS.SSF.load_table(wb.SSF);
	agregarFila(tablaPara(wb).tBodies[0],listaPara(wb));
	respuestasAlumnos[idTest(wb)].push(listaPara(wb));
}

var xlf = document.getElementById('xlf');
function handleFile(e) {
	var files = e.target.files;
	for (var i=0; i<files.length; i++)	{
		var f = files[i];
		var reader = new FileReader();
		var name = f.name;
		reader.onload = function(e) {
			var data = e.target.result;
			if(use_worker) {
				xw(data, process_wb);
			} else {
				var wb;
				if(rABS) {
					wb = X.read(data, {type: 'binary'});
				} else {
					var arr = fixdata(data);
					wb = X.read(btoa(arr), {type: 'base64'});
				}
				process_wb(wb);
			}
		};
		if(rABS) reader.readAsBinaryString(f);
		else reader.readAsArrayBuffer(f);
	};
	document.getElementById("botonExportar").disabled = false;
}

if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);