# PLCDataLog

Aplicação desktop WPF para **aquisição contínua de dados de PLC via Modbus TCP**, geração de histórico em CSV, envio automático de relatórios por e-mail, backup em pasta de rede e monitoramento/cópia de arquivos de receita.

Este projeto foi pensado para cenários industriais em que é necessário:
- acompanhar variáveis críticas de produção em tempo real;
- registrar alterações para rastreabilidade;
- distribuir automaticamente relatórios diários por e-mail;
- manter cópias de segurança em rede;
- sincronizar arquivos de receita entre origem local e destino de rede.

---

## Sumário

- [Visão geral](#visão-geral)
- [Principais funcionalidades](#principais-funcionalidades)
- [Arquitetura da solução](#arquitetura-da-solução)
- [Tecnologias](#tecnologias)
- [Requisitos](#requisitos)
- [Como executar](#como-executar)
- [Configuração da aplicação](#configuração-da-aplicação)
- [Guia de uso](#guia-de-uso)
  - [Dashboard](#dashboard)
  - [Diagnóstico](#diagnóstico)
  - [E-mail e automação](#e-mail-e-automação)
  - [Backup de rede](#backup-de-rede)
  - [Monitor de receitas](#monitor-de-receitas)
- [Formato de dados e logs](#formato-de-dados-e-logs)
- [Comportamento de agendamento diário](#comportamento-de-agendamento-diário)
- [Estrutura de projeto](#estrutura-de-projeto)
- [Desenvolvimento](#desenvolvimento)
- [Boas práticas de operação](#boas-práticas-de-operação)
- [Troubleshooting](#troubleshooting)
- [Roadmap sugerido](#roadmap-sugerido)
- [Licença](#licença)

---

## Visão geral

O **PLCDataLog** é um coletor de telemetria industrial com foco em simplicidade operacional. Ele mantém conexão com PLC via Modbus TCP, atualiza valores na interface, detecta mudanças e persiste snapshots em CSV diário.

Além da coleta, o sistema inclui recursos de automação para reduzir atividade manual:

- envio programado de CSV (inclusive lógica de “ontem”/catch-up);
- envio manual para teste operacional;
- alerta imediato de NOK quando há incremento no contador;
- cópia de logs para share de rede;
- monitoramento de pasta de receita com cópia automática e retry.

---

## Principais funcionalidades

### 1) Leitura de PLC (Modbus TCP)
- Polling contínuo configurável.
- Suporte a tipos:
  - `Word` (16 bits),
  - `DWord` (32 bits, com interpretação word swap),
  - `AsciiString` (swap por registrador),
  - `Coil` (com fallback de endereçamento para maior compatibilidade).

### 2) Registro CSV por alteração
- Criação de arquivo diário.
- Persistência de snapshots quando há mudança relevante.
- Histórico com timestamp para rastreabilidade.

### 3) E-mail SMTP
- Teste de SMTP pela interface.
- Envio manual do CSV do dia.
- Envio automático agendado com timezone configurável.
- Política de retry configurável.

### 4) Alerta imediato de NOK
- Identifica incremento de `Total_NOK`.
- Dispara e-mail de alerta com contexto operacional.

### 5) Backup de rede
- Cópia automática dos arquivos de log para pasta UNC/local configurada.
- Suporte a credenciais de rede.

### 6) Monitor de receitas
- `FileSystemWatcher` para detectar novos arquivos na origem.
- Cópia para destino de rede com retry e barra de progresso.

### 7) Descoberta de PLCs na rede
- Varredura de hosts em sub-rede `/24` na porta 502.
- Tentativa de identificação MAC e fabricante (OUI).

### 8) Execução em segundo plano (tray)
- Minimização para bandeja do sistema.
- Opção de mostrar/ocultar/sair pelo menu do ícone.

---

## Arquitetura da solução

O projeto segue uma estrutura simples de aplicação desktop com separação em:

- **UI (WPF):** interação do operador e exibição de estado.
- **Serviços:** persistência de settings, log CSV, SMTP, scheduler diário.
- **Modelos:** classes de configuração e contratos de dados.

Fluxo principal:

1. Operador configura conexão e parâmetros.
2. App inicia polling Modbus.
3. Valores são atualizados na grade.
4. Mudanças geram snapshots CSV.
5. Recursos opcionais (e-mail/backup/receitas) são executados conforme configuração.

---

## Tecnologias

- **.NET 9**
- **C# 13**
- **WPF** (desktop Windows)
- **FluentModbus** (cliente Modbus TCP)
- APIs nativas Windows para utilitários de rede (ARP)

---

## Requisitos

### Ambiente de execução
- Windows 10/11 ou Windows Server com Desktop Experience.
- Rede com acesso ao PLC (porta TCP 502).
- SMTP acessível (interno ou externo).

### Ambiente de desenvolvimento
- Visual Studio 2022/2026 com workload **.NET Desktop Development**.
- SDK do .NET 9 instalado.

---

## Como executar

### 1) Clonar

```bash
git clone https://github.com/Skyward-Games/PLCDataLog.git
cd PLCDataLog
```

### 2) Restaurar e compilar

```bash
dotnet restore
dotnet build -c Debug
```

### 3) Rodar

```bash
dotnet run --project PLCDataLog.csproj
```

> Também é possível abrir a solução no Visual Studio e executar com **F5**.

---

## Configuração da aplicação

Ao iniciar pela primeira vez, configure na tela de **Settings**:

### PLC
- **Host/IP** do PLC
- **Porta** (normalmente 502)
- **Poll interval** em ms
- **Unit ID**

### SMTP
- E-mail remetente
- Host/porta SMTP
- Usuário/senha
- Tipo de segurança
- Timeout e política de retry

### Destinatários
- Adicione e-mails válidos na lista.

### Automação
- Habilitar envio automático
- Hora/minuto
- Timezone
- Intervalo de retry

### Backup de rede
- Habilitar backup
- Pasta destino
- Credenciais (quando necessário)

### Monitor de receitas
- Habilitar monitoramento
- Pasta origem (watch)
- Pasta destino
- Credenciais de rede

---

## Guia de uso

## Dashboard

- **Start**: inicia conexão e polling.
- **Pause**: pausa leitura mantendo contexto para diagnóstico.
- **Stop**: encerra polling.
- Grade exibe tag, tipo, valor e horário de atualização.

Indicadores de status mostram:
- estado da conexão,
- endpoint ativo,
- mensagens de operação.

## Diagnóstico

Permite leitura sob demanda de uma tag para análise de payload bruto e decodificações alternativas.

Exibe:
- metadados da tag,
- bytes em hexadecimal,
- interpretação operacional,
- decodificação comparativa (BE/LE/word swap etc.).

## E-mail e automação

- **Test SMTP**: valida infraestrutura de envio.
- **Send Today CSV**: envio manual (não altera agenda automática).
- **Auto Send**: scheduler diário com timezone.
- **Reset Last Auto Send** e **Reset Yesterday CSV**: utilitários de recuperação operacional.

## Backup de rede

- Quando habilitado, cada atualização de CSV pode ser copiada para destino configurado.
- Botão de teste verifica leitura/escrita no destino.

## Monitor de receitas

- Observa criação/renomeação de arquivos na origem.
- Copia para destino com retentativa.
- Exibe progresso percentual.

---

## Formato de dados e logs

Os logs CSV são organizados por data (ano/mês/dia). Exemplo de árvore:

```text
DataRoot/
  Logs/
    2026/
      03/
        2026-03-21.csv
```

Também podem existir artefatos temporários para anexos de e-mail em:

```text
DataRoot/Temp/
```

---

## Comportamento de agendamento diário

O envio automático considera:

- timezone selecionado;
- data do último CSV enviado;
- lógica de “catch-up” para CSV de ontem quando aplicável;
- marcadores de controle persistidos em configuração.

Esse desenho reduz risco de perda de envio em reinicializações ou indisponibilidade temporária.

---

## Estrutura de projeto

Estrutura conceitual esperada:

- `MainWindow.xaml` / `MainWindow.xaml.cs`: UI principal e orquestração.
- `Models/`: classes de configuração e domínio.
- `Services/`: serviços de settings, CSV, SMTP, scheduler e integrações auxiliares.
- `App.xaml` / `App.xaml.cs`: bootstrap da aplicação.

---

## Desenvolvimento

### Compilação local

```bash
dotnet build -c Debug
```

### Publicação (exemplo)

```bash
dotnet publish -c Release -r win-x64 --self-contained false
```

### Pontos importantes no código

- Salvamento automático de settings com debounce para evitar gravação excessiva.
- Atualizações de UI via `Dispatcher`.
- Tratamento de falhas com status operacional e logs.
- Cuidado com credenciais (não registrar senha em logs).

---

## Boas práticas de operação

- Defina um host PLC estável (IP fixo ou DNS interno confiável).
- Ajuste `PollInterval` conforme capacidade do PLC e da rede.
- Configure destinatários e teste SMTP antes de habilitar automação.
- Valide permissões de escrita no destino de backup/receitas.
- Mantenha rotação e retenção de logs conforme política da planta.

---

## Troubleshooting

### Não conecta ao PLC
- Verifique IP/porta/unit ID.
- Teste conectividade de rede e firewall.
- Confirme se o PLC aceita Modbus TCP no endpoint configurado.

### E-mail não envia
- Confira host/porta/segurança SMTP.
- Verifique usuário/senha e bloqueios do provedor.
- Teste com envio manual antes de ativar agendamento.

### Backup de rede falhando
- Validar caminho UNC/local.
- Confirmar credenciais e permissão de escrita.
- Testar criação/exclusão de arquivo pelo botão de teste.

### Receita não copia
- Confirmar se watcher está habilitado e origem existe.
- Verificar destino e permissões.
- Revisar status exibido na UI para mensagens de retry.

---

## Roadmap sugerido

- Dashboard com gráficos e tendências por tag.
- Filtro de eventos e busca histórica.
- Exportação para banco de dados/OPC UA/REST.
- Health checks e telemetria estruturada.
- Empacotamento instalável com auto-update.

---

## Licença

Defina aqui o modelo de licença adotado pelo projeto (ex.: MIT, proprietário interno, etc.).

Se este software for utilizado em ambiente industrial produtivo, recomenda-se processo formal de validação, homologação e gestão de mudanças antes da entrada em operação.
