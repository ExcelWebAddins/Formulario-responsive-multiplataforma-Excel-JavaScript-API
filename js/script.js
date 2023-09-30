async function formulario(){
    let nombre = document.getElementById('nombre').value.trim();
    let apellido = document.getElementById('apellido').value.trim();
    let fecha = document.getElementById('fecha').value.trim();
    let correo = document.getElementById('correo').value.trim();
    let telefono = document.getElementById('telefono').value.trim();
    let contrasena = document.getElementById('contrasena').value.trim();
    let monto = document.getElementById('monto').value.trim();
    let pais = document.getElementById('pais').value.trim();
    let mensaje = document.getElementById('mensaje').value.trim();

    let soltero = document.getElementById('soltero').checked;
    let casado = document.getElementById('casado').checked;
    let estadoCivil = soltero === true ? 'Soltero' : 'Casado';
    let terminos = document.getElementById('terminos').checked;
    
    await Excel.run(async (context) => {
          let hoja = context.workbook.worksheets.getItem('Hoja1');
          let tabla = hoja.tables.getItem('Tabla1');
              hoja.protection.unprotect('123456');
          tabla.rows.add(
              null, // Añade filas al final de la tabla 
              [[nombre,apellido,fecha,correo,telefono,contrasena,monto,pais,estadoCivil,mensaje]],
              true // alwaysInsert en true especifica que las nuevas filas se inserten en la tabla.
          );   
          hoja.getUsedRange().format.autofitColumns();
          hoja.getUsedRange().format.autofitRows();
          hoja.protection.protect({allowAutoFilter:true}, '123456');
    await context.sync();
    });
    document.getElementById('limpiarFormulario').reset();
}




// Esta función toma un ID de elemento HTML como argumento y agrega un evento 'input'
// que permite que solo se ingresen caracteres de texto alfabèticos, espacios y tildes en el campo de entrada asociado

function soloTexto (elementId){
    //Obtener el elemento HTML con el ID proporcionado
    let elemento = document.getElementById(elementId);
    //Verificar si el elemento se encontrò en la pàgina
    if (elemento){
        //Agrega un evento 'input' al elemento encontrado
        elemento.addEventListener('input',function() {
        // Obtèn el valor actual del campo de entrada.
        let valor = this.value;
        // Utiliza una expresiòn regular que permita letras (alfabèticos), espacios y tildes.
        let valorFiltrado = valor.replace(/[^A-Za-záéíóúÁÉÍÓÚüÜ\sñÑ]/g , '');
        // Actualiza el valor del campo de entrada con el valor filtrado.
        this.value = valorFiltrado;
        });
    }
}



// Esta funciòn toma un ID de elemento HTML como argumento y agrega un evento 'input'
// que permite que solo se ingresen nùmeros en el campo de entrada asociado.

function soloNumero(elementId){
    //obtener el elemento HTML con el ID proporcionado
    let elemento = document.getElementById(elementId);
    //Verificar si el elemento se encontrò en la pàgina
    if (elemento){
        // Agregar un evento 'input' que se activarà cuando el usuario escriba en el campo de entrada
        elemento.addEventListener('input',function(){
          //Obtèn el valor actual del campo de entrada.
          let valor = this.value;
          //Utiliza una expresiòn regular para eliminar todos los caracteres que no sean dìgitos.
          let valorFiltrado = valor.replace(/[^0-9]/g , '');
          // Actualiza el valor del campo de entrada con el valor filtrado.
          this.value = valorFiltrado;
        });
    }
}



// Esta funciòn toma un ID de elemento HTML como argumento y le agrega un evento 'blur' 
// para formatear un nùmero ingresado en ese campo.

function formatoNumero(elementId){
   // Obtener el elemento HTML con el ID proporcionado
   const elemento = document.getElementById(elementId);
   // Verificar si el elemento se encontrò en la pàgina
   if (elemento){
   // Agregar un evento 'blur' que se activarà cuando el campo pierda el foco
   elemento.addEventListener('blur',function(){
   // Obtener el valor actual del campo de entrada
   let valor = this.value;
   // Eliminar caracteres no permitidos, dejando solo nùmeros , comas , puntos y guiones
      valor = valor.replace(/[^0-9.,-]/g , '');
   // Reemplazar comas por una cadena vacìa para que parseFloat pueda manejar el formato numèrico
     valor = valor.replace(/,/g , '');
   // Verificar si el valor es un nùmero vàlido y no està vacìo
    if (!isNaN(valor) && valor !== '' ){
        //Redondear el valor a dos decimales y luego formatearlo con comas para separar miles
        valor = parseFloat(valor).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
        // Verifica si el valor formateado es "0.00" y si es asì , establece el campo en blanco
           if (valor === "0.00"){
            this.value = '' ;
           } else{
            this.value = valor ;
           } 
    }else{
        //Si el valor no es vàlido , establece el campo en blanco
        this.value = '';
    }
   });
   }
}


// Llama a la funciòn y pasa el ID del elemento que deseas limitar a texto.
   soloTexto('nombre');
   soloTexto('apellido');
   soloTexto('mensaje');


// Llama a la funciòn y pasa el ID del elemento que deseas limitar a nùmeros.
   soloNumero('telefono');
   

// Llamar a la funciòn formatoNumero con el ID 'monto' para aplicar el formato a un campo especìfico en la pàgina
   formatoNumero('monto');   