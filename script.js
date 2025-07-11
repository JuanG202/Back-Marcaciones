

function enviarFormulario() {
  const nombre = document.getElementById('nombre').value.trim();
  const cedula = document.getElementById('cedula').value.trim();
  const agencia = document.getElementById('agencia').value.trim();
  const fechaEntrada = document.getElementById('horaEntrada').value;
  const horaSalida = document.getElementById('horaSalida').value;
  const observaciones = document.getElementById('observaciones').value.trim();

  if (!nombre || !cedula || !agencia) {
    alert("Por favor, completa los campos obligatorios (Nombre, Cédula y Agencia).");
    return;
  }

  // Deshabilitar el botón mientras se procesa
  const botonEnviar = document.querySelector('button[onclick="enviarFormulario()"]');
  botonEnviar.disabled = true;
  botonEnviar.textContent = 'Enviando...';

  const datos = {
    nombre,
    cedula,
    agencia,
    horaEntrada: fechaEntrada || '',
    horaSalida: horaSalida || '',
    observaciones
  };

  console.log('Enviando datos:', datos);

  fetch('https://back-marcaciones.vercel.app//registrar', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(datos)
  })
  .then(response => {
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    return response.json();
  })
  .then(data => {
    if (data.success) {
      let mensaje = data.message + "\n\n";
      mensaje += "Enlace al archivo Excel: " + data.urls.archivo;
      
      alert(mensaje);
      document.getElementById('formularioAseo').reset();
    } else {
      throw new Error(data.message || 'Error al guardar el registro');
    }
  })
  .catch(error => {
    console.error('Error:', error);
    alert("Error al enviar los datos: " + error.message);
  })
  .finally(() => {
    // Rehabilitar el botón
    botonEnviar.disabled = false;
    botonEnviar.textContent = 'Enviar';
  });
}

// Función para cargar y mostrar registros (opcional)
function cargarRegistros() {
  fetch('https://back-marcaciones.vercel.app//registros')
    .then(response => response.json())
    .then(data => {
      console.log('Registros cargados:', data);
      // Aquí puedes agregar código para mostrar los registros en tu página
    })
    .catch(error => console.error('Error al cargar registros:', error));
} 