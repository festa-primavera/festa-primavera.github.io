// =====================================================================
//  CONFIGURAÇÃO — Festa da Primavera 2026
//  E.M. Deputado Federal Ulisses Guimarães
// =====================================================================

// 1) Endereço do script do Google (Apps Script) ligado à planilha
//    "FESTA DA PRIMAVERA 2026 - VENDAS DE MESAS" no Google Drive.
//    É a URL que termina em /exec (Implantar → App da Web).
//    Enquanto ficar vazio, a página roda em MODO DEMONSTRAÇÃO.
export const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyKiqvBzYTu8CC9GJPcggqyviAW4k80kwrXjKqMYV_Npmyys34-x9quBCuiOQuZRkl0/exec";

// Intervalo, em segundos, para buscar novas vendas na planilha
export const INTERVALO_ATUALIZACAO = 4;

// 2) Valor de cada mesa em reais (já vem preenchido no formulário)
export const VALOR_MESA = 50;

// 3) Quantidade de mesas
export const TOTAL_MESAS = 50;

// 4) Vendedores
export const VENDEDORES = [
  "ALESSON MIRANDA PEREIRA",
  "MARINETE DA CONCEICAO TELES GONZAGA",
  "MARCIA MENEZES GOMES",
  "CARLOS ANDRE ALONSO DA SILVA",
];

// Quem recebe o dinheiro das vendas (prestação de contas). As vendas dele mesmo não entram.
export const TESOUREIRO = "CARLOS ANDRE ALONSO DA SILVA";

// 5) Séries e turmas
export const SERIES = [
  "1º Ano", "2º Ano", "3º Ano", "4º Ano", "5º Ano",
  "6º Ano", "7º Ano", "8º Ano", "9º Ano", "SE LIGA", "ACELERA",
];
export const TURMAS = ["A", "B", "C"];
export const PAGAMENTOS = ["Dinheiro", "PIX"];

// 6) Data da festa (horário de Manaus)
export const DATA_FESTA = "2026-11-12T00:00:00-04:00";
