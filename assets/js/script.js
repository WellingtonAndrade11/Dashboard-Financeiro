// Gráfico em Rosca
const graficoTorre = new Chart(document.getElementById("graficoTorre"), {
  type: "bar",

  data: {
    labels: ["Alimentos", "Bebidas", "Limpeza", "Tabacaria"],

    datasets: [
      // VALOR INVESTIDO
      {
        barPercentage: 0.7,
        categoryPercentage: 0.7,
        borderRadius: 8,

        label: "Valor Aplicado",

        data: [0, 0, 0, 0],

        backgroundColor: ["#4CAF50", "#2196F3", "#FFC107", "#F44336"],

        borderColor: ["#388E3C", "#1976D2", "#FFA000", "#D32F2F"],

        borderWidth: 1,
      },

      // PREVISÃO DE LUCRO
      {
        label: "Previsão de Lucro",

        data: [0, 0, 0, 0],

        backgroundColor: ["#81C784", "#64B5F6", "#FFD54F", "#E57373"],

        borderColor: ["#388E3C", "#1976D2", "#FFA000", "#D32F2F"],

        borderWidth: 1,
      },
    ],
  },

  options: {
    responsive: true,

    animation: {
      duration: 2000,
    },

    plugins: {
      datalabels: {
        anchor: "end",

        align: "top",

        formatter: (value) => {
          return value.toLocaleString("pt-BR", {
            style: "currency",
            currency: "BRL",
          });
        },

        font: {
          weight: "bold",
        },
      },
    },

    scales: {
      y: {
        beginAtZero: true,
      },
    },
  },

  plugins: [ChartDataLabels],
});

const graficoProdutos = new Chart(document.getElementById("graficoProdutos"), {
  type: "doughnut",
  data: {
    labels: ["Alimentos", "Bebidas", "Limpeza", "Tabacaria"],
    animation: {
      duration: 2000,
    },
    datasets: [
      {
        hoverOffset: 15,

        data: [0, 0, 0, 0],

        backgroundColor: ["#4CAF50", "#2196F3", "#FFC107", "#F44336"],

        borderWidth: 1,
      },
    ],
  },
  options: {
    responsive: true,
    plugins: {
      datalabels: {
        formatter: (value, context) => {
          let total = context.chart.data.datasets[0].data.reduce(
            (a, b) => a + b,
            0,
          );
          if (total <= 4) return "";
          let porcentagem = ((value / total) * 100).toFixed(1) + "%";
          return porcentagem;
        },
        color: "#fff",
        font: { weight: "bold", size: 14 },
      },
    },
  },
  plugins: [ChartDataLabels],
});

// Gráfico em Pizza
const graficoInvestido = new Chart(
  document.getElementById("graficoInvestido"),
  {
    type: "pie",
    data: {
      labels: ["Alimentos", "Bebidas", "Limpeza", "Tabacaria"],
      animation: {
        duration: 2000,
      },
      datasets: [
        {
          data: [0, 0, 0, 0],

          hoverOffset: 15,

          backgroundColor: ["#4CAF50", "#2196F3", "#FFC107", "#F44336"],

          borderWidth: 1,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        datalabels: {
          formatter: (value, context) => {
            let total = context.chart.data.datasets[0].data.reduce(
              (a, b) => a + b,
              0,
            );
            if (total <= 4) return "";
            return ((value / total) * 100).toFixed(1) + "%";
          },
        },
      },
    },
    plugins: [ChartDataLabels],
  },
);

document.getElementById("excelInput").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const dados = XLSX.utils.sheet_to_json(worksheet);

    const agora = new Date();

    const dataFormatada = agora.toLocaleDateString("pt-BR");

    const horaFormatada = agora.toLocaleTimeString("pt-BR", {
      hour: "2-digit",
      minute: "2-digit",
    });

    document.getElementById("ultimaAtualizacao").innerText =
      `${dataFormatada} às ${horaFormatada}`;

    let mensagem = document.getElementById("mensagemExcel");

    mensagem.classList.remove("d-none");

    setTimeout(() => {
      mensagem.classList.add("d-none");
    }, 3000);

    document
      .getElementById("btnAtualizar")
      .addEventListener("click", function () {
        location.reload();
      });

    document.querySelectorAll(".animar").forEach((elemento) => {
      elemento.classList.add("mostrar");
    });

    let investimentoCategoria = {
      Alimentos: 0,
      Bebidas: 0,
      Limpeza: 0,
      Tabacaria: 0,
    };

    let quantidadeCategoria = {
      Alimentos: 0,
      Bebidas: 0,
      Limpeza: 0,
      Tabacaria: 0,
    };

    let percentualLucro = {
      Alimentos: 0.1,
      Bebidas: 0.15,
      Limpeza: 0.1,
      Tabacaria: 0.05,
    };

    let totalProdutos = 0;
    let totalInvestido = 0;

    dados.forEach((item) => {
      let quantidade = Number(item["Quantidade"]) || 0;

      let precoUnitario =
        Number(String(item["Preço Unit."]).replace(",", ".")) || 0;

      let totalItem = quantidade * precoUnitario;

      // soma total de produtos
      totalProdutos += quantidade;

      // soma total investido geral
      totalInvestido += totalItem;
    });

    document.getElementById("totalProdutos").innerText = totalProdutos;

    document.getElementById("valorInvestido").innerText =
      totalInvestido.toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      });

    // PRIMEIRO: calcular os dados

    dados.forEach((item) => {
      let categoria = item["Categoria"];
      let quantidade = Number(item["Quantidade"]) || 0;
      let precoUnitario =
        Number(String(item["Preço Unit."]).replace(",", ".")) || 0;

      // cálculo correto
      let totalInvestidoDoProduto = quantidade * precoUnitario;

      if (quantidadeCategoria[categoria] !== undefined) {
        quantidadeCategoria[categoria] += quantidade;
      }

      if (investimentoCategoria[categoria] !== undefined) {
        investimentoCategoria[categoria] += totalInvestidoDoProduto;
      }
    });

    // SEGUNDO: montar a tabela

    let tabela = document.getElementById("tabelaEstoque");
    tabela.innerHTML = "";

    Object.keys(quantidadeCategoria).forEach((categoria) => {
      let quantidade = quantidadeCategoria[categoria];
      let investido = investimentoCategoria[categoria];

      let percentual = percentualLucro[categoria] || 0;
      let lucro = investido * percentual;

      tabela.innerHTML += `
    <tr>
      <td>${categoria}</td>
      <td>${quantidade}</td>
      <td>${investido.toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      })}</td>
      <td>${lucro.toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL",
      })}</td>
    </tr>
  `;
    });

    let lucroCategoria = {};

    // calcula o lucro de cada categoria
    Object.keys(investimentoCategoria).forEach((categoria) => {
      let investido = investimentoCategoria[categoria];

      let percentual = percentualLucro[categoria] || 0;

      lucroCategoria[categoria] = investido * percentual;
    });

    Object.keys(investimentoCategoria).forEach((categoria) => {
      let investido = investimentoCategoria[categoria];
      let percentual = percentualLucro[categoria] || 0;

      let lucro = investido * percentual;

      lucroCategoria[categoria] = lucro;
    });

    graficoTorre.data.datasets[0].data = Object.values(investimentoCategoria);

    graficoTorre.data.datasets[1].data = Object.values(lucroCategoria);

    graficoTorre.update();

    graficoProdutos.data.datasets[0].data = Object.values(quantidadeCategoria);

    graficoProdutos.update();

    graficoInvestido.data.datasets[0].data = Object.values(
      investimentoCategoria,
    );

    graficoInvestido.update();
  };

  reader.readAsArrayBuffer(file);
});
