// Recebe o CSV enviado pelo Google Apps Script (rotina semanal) e guarda no Vercel Blob,
// sob uma chave fixa — a versão mais recente sempre sobrescreve a anterior.
//
// Variáveis de ambiente necessárias no Vercel (Settings → Environment Variables):
//   WEBHOOK_SECRET        — segredo compartilhado com o Apps Script (você escolhe o valor)
//   BLOB_READ_WRITE_TOKEN — criado automaticamente ao ativar o Vercel Blob no projeto
//
// O Apps Script deve enviar um POST com o CSV no corpo (texto puro) e o header:
//   x-webhook-secret: <mesmo valor de WEBHOOK_SECRET>

import { put } from "@vercel/blob";

const CHAVE_ARQUIVO = "venda-salvados-atual.csv";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Método não permitido. Use POST." });
    return;
  }

  const segredoRecebido = req.headers["x-webhook-secret"];
  if (!segredoRecebido || segredoRecebido !== process.env.WEBHOOK_SECRET) {
    res.status(401).json({ error: "Não autorizado — segredo ausente ou incorreto." });
    return;
  }

  try {
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const csvText = Buffer.concat(chunks).toString("utf-8");

    if (!csvText || csvText.trim().length < 10) {
      res.status(400).json({ error: "Corpo vazio ou inválido." });
      return;
    }

    const blob = await put(CHAVE_ARQUIVO, csvText, {
      access: "public",
      addRandomSuffix: false,
      allowOverwrite: true,
      contentType: "text/csv; charset=utf-8",
    });

    res.status(200).json({
      ok: true,
      recebidoEm: new Date().toISOString(),
      tamanhoBytes: Buffer.byteLength(csvText, "utf-8"),
      url: blob.url,
    });
  } catch (err) {
    res.status(500).json({ error: "Erro ao processar o arquivo: " + err.message });
  }
}
