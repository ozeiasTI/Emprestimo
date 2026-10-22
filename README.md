# Empréstimo

Aplicação desktop em Python para administrar categorias e itens de empréstimo. O SQLite fica no computador host e é compartilhado entre contas do Windows nesse host.

## Requisitos

- Python 3.10 ou superior, com Tkinter instalado.
- Dependências da aplicação listadas em `requirements.txt` (inclui Pillow para exibir `foto.webp`).

## Instalação e execução

Na pasta do projeto, instale as dependências e inicie a aplicação:

```powershell
python -m pip install -r requirements.txt
python Emprestimo.py
```

## Recursos desta versão

- Cadastro, consulta, edição, exclusão e pesquisa de categorias.
- Cadastro, consulta, edição, exclusão e pesquisa de itens por categoria.
- Status de item: Disponível (padrão), Emprestado ou Em manutenção.
- Cadastro manual de usuários com nome, matrícula, e-mail e tipo (Servidor, Externo ou Colaborador).
- Edição e pesquisa de usuários, filtro por tipo e exportação da lista filtrada para Excel e PDF.
- Importação de servidores a partir de `.xls` ou `.xlsx`; sincroniza pela matrícula e, quando ela muda, tenta preservar a identidade usando e-mail ou nome único. Empréstimos e mensagens de e-mail mantêm seus vínculos com o mesmo usuário.
- A tela de usuários orienta a exportação de servidores no SUAP IFRO (Gestão de Pessoas > Servidores > campus e Ativo > Pesquisar > Exportar XLS).
- A ordem/posição da linha na planilha não identifica o servidor. Servidores ausentes sem histórico são removidos; os que têm empréstimos ou e-mails ficam preservados. Registros cuja identidade não pode ser associada com segurança são mantidos separadamente para evitar mover histórico para outra pessoa.
- Envio do histórico completo de empréstimos de um usuário selecionado em uma única mensagem com anexos Excel e PDF.
- Visualização dos itens em lista ou grade e filtro por categoria.
- O número do item é único dentro da categoria e pode se repetir em categorias diferentes.
- Exportação das listas filtradas de categorias e itens para PDF e Excel.
- Os relatórios PDF são gerados em orientação paisagem, com colunas ampliadas para acomodar nomes longos de itens.
- O seletor de empréstimos mantém itens marcados ao alternar categorias e exibe a lista completa selecionada antes de registrar a retirada.
- A janela principal é aberta maximizada após a tela de abertura.
- Painel inicial com indicadores, acervo por categoria em coluna compacta, itens recentes em área ampliada com filtros por categoria, nome e número, e resumo de enviados, pendentes e falhos. O gráfico mensal foi removido.
- Ao iniciar, o sistema mostra os empréstimos em aberto e permite enviar uma cobrança agrupada de todos os itens de uma pessoa.
- Janela ampla de seleção de usuário para empréstimos, com busca digitável por nome, matrícula, e-mail ou tipo.
- Na devolução, a lista apresenta todos os itens em aberto da pessoa selecionada; é possível selecionar vários itens específicos ou todos, gerando uma única notificação.
- Durante o envio do histórico por e-mail, uma janela de processamento bloqueia ações duplicadas até a conclusão.
- A logo `logo.png` aparece no ícone do aplicativo e na navegação lateral.
- Tela de abertura com a foto `foto.webp` e página **Sobre** com biografia, formação, experiência e contatos do autor.
- Temas predefinidos Verde, Azul, Roxo, Grafite, Turquesa, Vinho, Âmbar, Oliva, Ciano, Rosa, Laranja, Noturno, Floresta e Lavanda na aba **Cores** das configurações; aplicar um tema reinicia o aplicativo.
- Campos de texto com exemplos/indicações dos dados esperados.
- Configuração do local do arquivo SQLite no menu **Configurações**. O padrão e as preferências ficam em `%PUBLIC%\Documents\Emprestimo`, compartilhados entre contas do host.
- Aba **Backup** nas configurações: por padrão cria a pasta `backup` dentro da pasta do banco ativo (por exemplo, `Downloads\backup` se o banco estiver em Downloads). É possível escolher outro local raiz; o sistema cria a subpasta `backup` e exibe o caminho final. Faça uma cópia completa sob demanda ou programe backups diários/semanais. A rotina verifica o vencimento ao abrir o app e enquanto ele permanece em execução; se o app ficar fechado na data prevista, o backup vencido é feito na próxima abertura.
- Importação de um banco SQLite pela aba **Backup**: o arquivo é validado antes de substituir os dados atuais. Não há backup automático antes da importação; faça um backup manual primeiro caso queira preservar a versão atual.
- Cadastro de e-mail, senha de app, instituição, remetente, servidor/porta SMTP e TLS na aba **E-mail** das configurações.
- O servidor SMTP padrão é `smtp.gmail.com`, porta `587`, com TLS habilitado.
- Assinatura de e-mail editável na aba **E-mail**, aplicada às notificações de retirada, devolução e cobrança.
- A senha SMTP é armazenada no banco do host para permitir o uso compartilhado pela aplicação; não há criptografia local por perfil do Windows. Senhas salvas por versões antigas precisam ser informadas novamente uma vez.
- Retirada de vários itens para um usuário em uma única operação, com uma única confirmação por e-mail contendo uma tabela completa.
- Devolução de vários itens do mesmo usuário em uma única operação, com um único e-mail listando os itens devolvidos.
- Cobranças agrupadas por usuário: o lembrete lista todos os empréstimos ainda em aberto daquele usuário.
- Aba **Fila de e-mails** com estado visual de enviados, na fila, em envio e com falha; duplo clique mostra detalhes e o motivo do erro, com reenvio de mensagens pendentes/com falha.
- Extrato completo de empréstimos em PDF e Excel, com usuário, datas, situação e estado das notificações.
- Duplo clique nos itens recentes do painel mostra quem está com o item quando houver empréstimo aberto.
- Os estados de itens e devoluções nas listas, no painel e no histórico usam ícones compactos, com legenda para disponibilidade, empréstimo, manutenção e devolução.
- Para vários computadores, execute a aplicação no host e permita o acesso por sessão remota. Não abra o mesmo arquivo SQLite diretamente por uma pasta de rede.
- Ao executar o script Python, o caminho do banco e o tema ficam nas configurações compartilhadas do host.
- Uma categoria que possui itens não pode ser excluída até que esses itens sejam removidos ou transferidos.

## Gerar o executável no Windows

Com o PyInstaller instalado no ambiente Python, execute na pasta do projeto:

```powershell
pyinstaller --onefile --windowed --name "Emprestimo" --icon "logo.ico" --hidden-import "PIL._imagingtk" --add-data "logo.png;." --add-data "foto.webp;." --add-data "logo.ico;." Emprestimo.py
```

O executável será criado em `dist\Emprestimo.exe`; as imagens e o ícone são empacotados. Para compilar em outra máquina, instale as dependências do projeto e o PyInstaller nesse ambiente.

Esse comando gera um executável para Windows na arquitetura do ambiente de compilação. Copiar o `.exe` para outro computador não sincroniza os dados: no modo executável, cada cópia mantém seu próprio banco ao lado do `.exe`. Para compartilhar dados entre computadores, execute o aplicativo no host e acesse-o por sessão remota; não use o mesmo arquivo SQLite diretamente em uma pasta de rede.

No modo executável (`--onefile`), o banco `emprestimo.db`, as configurações `config.json` e a pasta `backup` são criados junto ao `.exe`. Na primeira execução, se houver uma configuração antiga do aplicativo neste computador, o sistema tenta copiar o banco antigo para essa pasta antes de iniciar, sem remover a cópia original. Coloque o executável em uma pasta com permissão de gravação, como a Área de Trabalho; pastas protegidas pelo Windows, como `Arquivos de Programas`, podem impedir a criação desses arquivos. Ao executar `Emprestimo.py` diretamente, o banco e as configurações continuam usando a pasta compartilhada em `%PUBLIC%\Documents\Emprestimo`.