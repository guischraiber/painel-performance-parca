// Devolve o CSV mais recente recebido via /api/receber-dados, para o painel carregar
// automaticamente ao abrir. Se ainda não houver nenhum arquivo recebido, devolve 404 —
// o painel então cai no upload manual.

import { list } from "@vercel/blob";

const CHAVE_ARQUIVO = "venda-salvados-atual.csv";

export default async function handler(req, res) {
  try {
    const { blobs } = await list({ prefix: CHAVE_ARQUIVO });
    if (!blobs || blobs.length === 0) {
      res.status(404).json({ error: "Nenhum arquivo automático disponível ainda." });
      return;
    }

    const blob = blobs[0];
    const fileRes = await fetch(blob.url);
    if (!fileRes.ok) {
      res.status(502).json({ error: "Falha ao buscar o arquivo armazenado." });
      return;
    }
    const text = await fileRes.text();

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("X-Atualizado-Em", blob.uploadedAt);
    res.status(200).send(text);
  } catch (err) {
    res.status(500).json({ error: "Erro ao buscar dados: " + err.message });
  }
}
