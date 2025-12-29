# Configurador de Respostas Automáticas (Out of Office)

Script PowerShell para configurar respostas automáticas de ausência no Microsoft 365 utilizando a Graph API.

## 📋 Descrição

Este script permite configurar automaticamente mensagens de resposta automática (Out of Office/OOO) para um usuário do Microsoft 365. É ideal para:

- Configurar mensagens de férias
- Definir períodos de ausência
- Redirecionar demandas para um substituto
- Definir datas e horários automáticos de ativação e desativação

## 🔧 Requisitos

- **Windows PowerShell 5.0+** ou **PowerShell Core 7.0+**
- Acesso à conta de administrador do Microsoft 365
- Credenciais de aplicação Azure AD:
  - Tenant ID
  - Client ID (Application ID)
  - Client Secret

## 🚀 Como Usar

### 1. Preparar Credenciais Azure AD

1. Acesse o [Azure Portal](https://portal.azure.com)
2. Navegue até **Azure Active Directory** > **Registros de aplicativo**
3. Crie um novo registro de aplicativo
4. Configure as permissões necessárias:
   - `User.Read.All`
   - `Mail.ReadWrite`
5. Crie um segredo do cliente
6. Anote: **Tenant ID**, **Client ID** e **Client Secret**

### 2. Configurar o Script

Abra o arquivo PowerShell e substitua as variáveis no topo:

```powershell
$TenantId     = "Seu-Tenant-ID-Aqui"
$ClientId     = "Seu-Client-ID-Aqui"
$ClientSecret = "Seu-Client-Secret-Aqui"
```

### 3. Executar o Script

```powershell
.\script_ooo.ps1
```

### 4. Fornecer Informações Solicitadas

O script solicitará as seguintes informações:

1. **Nome completo do usuário** - Ex: João Silva
2. **E-mail do usuário** - Ex: joao.silva@empresa.com
3. **Data e hora de início** - Formato: `YYYY-MM-DD HH:MM` (Ex: 2025-12-29 09:00)
4. **Data e hora de fim** - Formato: `YYYY-MM-DD HH:MM` (Ex: 2025-12-30 18:00)
5. **Redirecionar?** - Digite `S` para sim ou `N` para não
6. *(Opcional)* **Nome do substituto** - Se redirecionar
7. *(Opcional)* **E-mail do substituto** - Se redirecionar

## 📝 Exemplo de Uso

```
1. Nome completo do usuário: Maria Santos
2. E-mail do usuário: maria.santos@empresa.com
3. Data e hora de início (ex: 2025-12-29 09:00): 2025-12-29 09:00
4. Data e hora de fim (ex: 2025-12-30 18:00): 2026-01-10 18:00
5. Redirecionar? (S/N): S
6. Nome do substituto: Pedro Costa
7. E-mail do substituto: pedro.costa@empresa.com
```

## 💬 Mensagem Gerada

A mensagem automática será formatada como:

```
Prezados(as),

Informo que estarei de férias no período de 29/12/2025 a 10/01/2026, 
retornando às minhas atividades no dia 11/01/2026.

Durante esse período, não estarei acompanhando as demandas. Para assuntos 
relacionados à minha área de atuação, gentileza direcionar para Pedro Costa, 
no e-mail pedro.costa@empresa.com, que poderá auxiliá-los no que for necessário.

Agradeço pela compreensão.

Atenciosamente,
Maria Santos
```

## 🌍 Fuso Horário

O script utiliza por padrão o fuso horário: **E. South America Standard Time** (Brasília)

Para modificar, altere a variável:
```powershell
$TimeZone = "E. South America Standard Time"
```

[Ver lista completa de fusos horários suportados](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones)

## ✅ Validações

O script valida:

- ✓ Formato das datas de início e fim
- ✓ Data de início deve ser anterior à data de fim
- ✓ Autenticação com Azure AD
- ✓ Resposta da Graph API

## ⚠️ Possíveis Erros

| Erro | Solução |
|------|---------|
| "Acesso recusado (401)" | Verifique credenciais (Tenant ID, Client ID, Client Secret) |
| "Recurso não encontrado (404)" | Verifique se o e-mail do usuário está correto |
| "Formato inválido" | Use o formato de data correto: `YYYY-MM-DD HH:MM` |
| "Início deve ser antes do fim" | Data de início deve ser anterior à data de fim |

## 🔒 Segurança

⚠️ **Importantes:**
- Nunca compartilhe o `Client Secret`
- Use variáveis de ambiente para armazenar credenciais em produção
- Revise as permissões do aplicativo Azure AD regularmente
- Mantenha o script em local seguro com controle de acesso

## 📚 Recursos Adicionais

- [Documentação Microsoft Graph - Mailbox Settings](https://learn.microsoft.com/en-us/graph/api/user-update-mailboxsettings)
- [Azure AD Application Registration](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [PowerShell Invoke-RestMethod](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-restmethod)

## 📄 Licença

Este script é fornecido como está, sem garantias.

---

**Versão:** 1.0  
**Data de Atualização:** Dezembro de 2025
