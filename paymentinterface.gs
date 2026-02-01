<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; }
    .container { width: 600px; margin: 20px auto; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th, td { padding: 10px; border: 1px solid #ddd; text-align: left; }
    th { background-color: #f4f4f4; }
    input, button { padding: 10px; margin-top: 10px; width: 100%; }
  </style>
</head>
<body>
  <div class="container">
    <h2>Registro de Pagamentos</h2>
    <div id="paymentData">
      <!-- Dados dos pagamentos serão carregados aqui -->
    </div>
    <h3>Registrar Pagamento</h3>
    <form id="paymentForm">
      <label for="uniqueID">Unique ID:</label>
      <input type="text" id="uniqueID" name="uniqueID">
      
      <label for="paymentAmount">Valor Pago:</label>
      <input type="number" id="paymentAmount" name="paymentAmount" step="0.01">
      
      <label for="paymentDate">Data de Pagamento:</label>
      <input type="date" id="paymentDate" name="paymentDate">
      
      <label for="parcela">Número da Parcela:</label>
      <input type="number" id="parcela" name="parcela">
      
      <button type="button" onclick="submitPayment()">Registrar Pagamento</button>
    </form>
  </div>
  
  <script>
    function submitPayment() {
      const uniqueID = document.getElementById('uniqueID').value;
      const paymentAmount = parseFloat(document.getElementById('paymentAmount').value);
      const paymentDate = document.getElementById('paymentDate').value;
      const parcela = parseInt(document.getElementById('parcela').value);

      google.script.run.recordPayment(uniqueID, paymentAmount, paymentDate, parcela);
    }
    
    function loadPaymentData() {
      google.script.run.withSuccessHandler(function(data) {
        const container = document.getElementById('paymentData');
        container.innerHTML = '<table><tr><th>Unique ID</th><th>Nome Completo</th><th>Valor Total</th><th>Status</th></tr>';
        data.forEach(row => {
          container.innerHTML += `<tr><td>${row[0]}</td><td>${row[1]}</td><td>${row[5]}</td><td>${row[9]}</td></tr>`;
        });
        container.innerHTML += '</table>';
      }).getPaymentData();
    }

    loadPaymentData();
  </script>
</body>
</html>
