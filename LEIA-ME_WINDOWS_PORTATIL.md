# Sistema de Cobrança — modo portátil para Windows

## Onde ficam os dados

Mantenha `SistemaCobranca.exe` e `MODO_PORTATIL.txt` juntos, na mesma pasta do pendrive. Na primeira abertura, o programa cria `dados/sistema_cobranca.db` nessa pasta. Todos os computadores que abrirem **essa mesma cópia do pendrive** verão o mesmo banco.

Não remova o pendrive enquanto o sistema estiver aberto. Feche o programa, espere alguns segundos e use a opção do Windows para ejetar a unidade com segurança. Não abra o mesmo banco em dois computadores ao mesmo tempo e evite serviços de sincronização de arquivos para a pasta `dados`.

Depois de fechar o programa, copie a pasta `dados` inteira periodicamente para um local seguro. Os dados dos devedores são sensíveis: proteja o pendrive e a cópia de segurança contra acesso não autorizado. O programa não oferece criptografia do banco.

## Primeiro uso, quando já existe um banco no computador

Se a pasta `dados` do pendrive ainda não tem banco e o programa encontrar um banco local naquele computador, ele perguntará se deve **copiar** esse banco para o pendrive. Escolha **Sim** para levar os dados existentes. O banco original permanece no computador. Escolha **Não** apenas se quiser começar vazio; **Cancelar** fecha o programa sem alterar os dados.

Se o pendrive já tem um banco, o programa usa esse banco e **não mistura automaticamente** os dados locais de outros computadores. Caso existam bancos diferentes em mais de um computador, não escolha um deles sem antes conferir qual é o mais completo: a junção desses bancos exige um procedimento específico.

## Como gerar o executável em um computador Windows

Se você já usa GitHub Actions, envie os arquivos deste pacote à raiz do repositório, incluindo `.github/workflows/gerar-windows.yml`. Na aba **Actions**, execute manualmente **Gerar aplicativo Windows**. O artefato `SistemaCobranca-Windows-portatil` conterá um ZIP com `SistemaCobranca.exe` e `MODO_PORTATIL.txt` na mesma pasta. Ele não contém nenhum banco de dados pessoal.

### Alternativa: compilar diretamente no Windows

Extraia o pacote-fonte em uma pasta do Windows. Com Python instalado nesse computador, execute `EMPACOTAR_WINDOWS.bat`. Ele instala as dependências de `requirements.txt` e o PyInstaller, compila o programa e gera `SistemaCobranca_Carla_portatil.zip`.

Se preferir fazer a compilação manualmente, o comando é:

```
py -m PyInstaller --noconfirm --clean SistemaCobranca_Windows.spec
```

Na compilação manual, copie para uma pasta no pendrive:

- `dist/SistemaCobranca.exe`
- `MODO_PORTATIL.txt`

O ZIP gerado pelo script já contém os dois arquivos necessários. Ao abrir, o programa cria a pasta `dados` automaticamente. Em atualizações futuras, substitua **somente o executável**: preserve `MODO_PORTATIL.txt` e a pasta `dados`.

O executável deve ser gerado no Windows; o PyInstaller no Mac não gera um `.exe` de Windows. Teste a primeira cópia em um pendrive de teste antes de entregar à Carla.
