# Sistema de Cobrança — nova organização para teste no Mac

Esta pasta reúne os arquivos do aplicativo desktop. O parser estável não foi alterado.

1. Faça uma cópia de segurança da pasta atual do sistema.
2. Extraia os arquivos do pacote sobre essa pasta, substituindo os arquivos de mesmo nome.
3. Feche e abra novamente `app_desktop.py` do mesmo modo que já usa no Mac.
4. O sistema abre em **Importar / salvar**. A data da última carga aparece nessa aba. As ações estão separadas em três áreas: **Exportar Base Mestre** (Excel e CSV), **Exportar comunicados** e **Importar telefones** (planilha e modelo).
5. Abra **Condomínios / devedores**, selecione um condomínio e um devedor. Os contatos aparecem no painel inferior esquerdo, ao mesmo tempo que os débitos aparecem na ficha à direita.
   A lista de devedores mostra economia, nome, quantidade de débitos e total; os telefones ficam apenas no painel de contatos.
6. Para adicionar um telefone ou e-mail, clique em **Novo**, escolha o tipo, informe o valor e salve. Para corrigir um contato, selecione-o na lista e edite os campos. É possível manter vários telefones e e-mails por devedor.
7. Ao registrar um contato, escolha o canal e confira o número ou e-mail usado. Esse valor fica preservado no histórico, mesmo se o cadastro for editado depois.
8. Para uma nova carga, volte à primeira aba, selecione os PDFs e processe-os.

No Mac, o aplicativo abre em tela cheia para deixar o painel de contatos visível. Pressione **Esc** para voltar ao tamanho normal. O painel de contatos começa na mesma altura da ficha do devedor; as tabelas têm barras de rolagem quando necessário.

Na primeira aba também ficam Base Mestre CSV, importação de telefones e modelo de planilha. Planilhas de importação podem conter mais de uma linha de telefone para o mesmo devedor.

O comunicado individual segue a regra atual do modelo: só inclui parcelas vencidas há mais de 30 dias. O sistema não envia mensagens. O histórico permite registrar contatos realizados manualmente por telefone, e-mail, SMS, carta ou outro canal. Registros antigos permanecem no banco, inclusive os de canais que não aparecem mais como opção para novos contatos.

As abas dos Excel de comunicados — tanto no pacote por condomínio quanto no arquivo individual — agora identificam o devedor e a economia. O Excel limita o nome da aba a 31 caracteres; em nomes longos, o nome do devedor é abreviado na aba, mas aparece completo dentro do comunicado.

O banco de dados existente permanece em `~/Library/Application Support/SistemaCobranca/sistema_cobranca.db`, fora da pasta do aplicativo. Substituir os arquivos do programa não o apaga. Na primeira abertura desta versão, o banco é atualizado para guardar múltiplos contatos; uma cópia anterior é criada no mesmo diretório com o nome `sistema_cobranca.antes_v4.db`.
