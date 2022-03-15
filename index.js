// Generador de boletos aleatorios de 1 a 1000
// se generan siempre boletos cambiando unos parametros iniciales 

// Require library
var xl = require('excel4node');

console.log("Bienvenido al juego de boletos aleatorios");
let num_boletos = 300;
let id_proyecto = 1;
let nom_proyecto = "dia de las madres 2022";

longitud_proyecto = longitud(nom_proyecto);
boletos = generador_boletos(num_boletos, id_proyecto, longitud_proyecto);
console.log(boletos);

// AGREGAR A EXCEL
agregar_excel(boletos,num_boletos);

/*
    ********************************************************************
                        FUNCIONES
    ********************************************************************
*/
function longitud(cadena) {
    let longitud = cadena.length;
    return longitud;
}

function generador_boletos(num_boletos, id_proyecto, longitud_proyecto){
    let boletos_generados = [];
    let boleto_aleatorio
    let boleto_aleatorio_completo
    for (var i = 0; i <= (num_boletos + 1); i++) {
        //i++;
        boleto_aleatorio = (id_proyecto * longitud_proyecto + longitud_proyecto - id_proyecto) * i;
        // convertir el numero a string
        boleto_aleatorio_completo = boleto_aleatorio.toString();
        // llenado de ceros
        switch (longitud(boleto_aleatorio_completo)){
            case 1:
                boleto_aleatorio_completo = "00000" + boleto_aleatorio_completo
                break;
            case 2:
                boleto_aleatorio_completo = "0000" + boleto_aleatorio_completo
                break
            case 3:
                boleto_aleatorio_completo = "000" + boleto_aleatorio_completo
                break
            case 4:
                boleto_aleatorio_completo = "00" + boleto_aleatorio_completo
                break  
            case 5:
                boleto_aleatorio_completo = "0" + boleto_aleatorio_completo
                break  
            default:
                break;
        }
        boletos_generados.push(boleto_aleatorio_completo);
    }
    return boletos_generados
}

// FUNCION DE EXCEL
function agregar_excel(datos, num_boletos){ 
    // Create a new instance of a Workbook class
    var wb = new xl.Workbook();
    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');
    //var ws2 = wb.addWorksheet('Sheet 2');
    
    // Create a reusable style
    var style = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    },
    //numberFormat: '#,##0.00; ($#,##0.00); -',
    });
    
    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    ws.cell(1, 1)
    .string("Boleto")
    .style(style);
    
    ws.cell(1, 2)
    .string("ID aleatorio")
    .style(style);
    
    for (var i = 1; i <= (num_boletos); i++) {
        ws.cell((i + 1), 1)
        .number(i)
        .style(style);
    }
    
    for (var i = 1; i <= (num_boletos); i++) {
    ws.cell((i + 1), 2)
    .string(datos[i])
    .style(style);
    }
    
    wb.write('Excel.xlsx');
    
}