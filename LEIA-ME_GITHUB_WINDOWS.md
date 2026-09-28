# Sistema de Cobrança — fontes completas para Windows

Este ZIP contém o código completo da versão de teste com telefones extraídos dos
PDFs, os modelos Excel, `requirements.txt` e os dois arquivos `.spec`.

Para atualizar o GitHub, envie os arquivos deste ZIP **à raiz do repositório**,
substituindo os homônimos. Não envie a pasta que contém o ZIP como uma subpasta.
Preserve o caminho `.github/workflows/gerar-windows.yml` dentro do repositório:
esse é o workflow incluído para gerar o Windows. Na aba **Actions**, escolha
**Gerar Windows portátil (versão completa)** e execute-o manualmente.
Se existirem workflows Windows antigos, não os execute por engano; eles podem
gerar versões anteriores do programa.

O código usa `SistemaCobranca_Windows.spec` para um executável único. O arquivo
`SistemaCobranca.spec` também está incluído caso seu workflow use esse nome.
Ambos incluem os módulos de leitura de telefones.

Se a Carla usa o sistema na pendrive, mantenha `MODO_PORTATIL.txt` ao lado de
`SistemaCobranca.exe` depois de baixar o artefato. O banco fica na pasta
`dados` ao lado deles. **Nunca envie `dados/` nem `sistema_cobranca.db` ao GitHub.**

Nesta versão, a coluna de observações dos contatos tem largura 390. As notas
de telefone sem DDD ou com formato atípico aparecem de forma abreviada; a
recuperação de dados reconhece tanto as notas antigas quanto as novas.
