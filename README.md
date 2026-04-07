# 📊 Kobo Dashboard — 27ª Assembleia Distrital Matola

Google Apps Script que importa submissões do KoboToolbox e gera um dashboard dinâmico no Google Sheets.

---

## ✨ Funcionalidades

| Funcionalidade | Detalhe |
|---|---|
| **Importação paginada** | Busca todas as submissões via KoboToolbox v2 API com suporte a paginação |
| **Normalização de campos** | Remove prefixos de grupo automáticamente (`grupo/campo` → `campo`) |
| **Aba `Raw_Kobo`** | Dados limpos de todas as submissões |
| **Aba `Kobo_Meta`** | Data/hora da última sync, total de respostas, Asset UID |
| **Dashboard dinâmico** | Gerado via Apps Script com formatação visual completa |
| **Filtros** | Por perfil, género e proveniência/igreja |

### Dashboard inclui:
- 🔢 **KPI cards** — total respostas, média global, % espaço de opinião, melhor indicador
- 📋 **Avaliação por indicador** — barras de progresso + badge de classificação (1–5)
- 📊 **Distribuição** — tabela Excelente/Bom/Satisfatório/Mau/Muito Mau por indicador
- 👥 **Perfil dos participantes** — por perfil, género e igreja
- 🗣️ **Espaço de opinião** — global e por perfil
- 💬 **Comentários abertos** — tabela com pontos fortes e áreas de melhoria

---

## 🚀 Instalação

### 1. Pré-requisitos
- Conta Google (Google Sheets + Apps Script)
- Token de acesso KoboToolbox
- Asset UID do formulário no KoboToolbox

### 2. Configurar o script

Abre o ficheiro `KoboDashboard.gs` e edita a secção `CONFIG` no topo:

```javascript
const CONFIG = {
  ASSET_UID : 'SEU_ASSET_UID_AQUI',
  TOKEN     : 'SEU_TOKEN_AQUI',
  BASE_URL  : 'https://eu.kobotoolbox.org/api/v2',  // ou https://kf.kobotoolbox.org/api/v2
  PAGE_SIZE : 30000,
  EVENT_NAME: 'Nome do Evento',
};
```

> ⚠️ **Segurança:** Nunca comites o token real no GitHub. Usa o ficheiro `.env.example` como referência e guarda as credenciais nas [Properties do Apps Script](https://developers.google.com/apps-script/guides/properties).

### 3. Instalar no Google Sheets

1. Abre um **Google Sheets novo** (ou existente)
2. Vai a **Extensions → Apps Script**
3. Apaga o código existente
4. Cola o conteúdo de `KoboDashboard.gs`
5. Guarda com `Ctrl+S`

### 4. Executar pela primeira vez

1. No editor Apps Script, selecciona `syncKobo` no dropdown de funções
2. Clica **▶ Run**
3. Aceita as permissões que o Google pede (necessário apenas na primeira vez)
4. Aguarda a conclusão — verás toasts no Sheets

### 5. Uso normal (depois da instalação)

O menu **📊 Kobo Dashboard** aparece automaticamente no Google Sheets:

| Opção | Descrição |
|---|---|
| 🔄 Sincronizar dados do KoboToolbox | Busca novas submissões e reconstrói o dashboard |
| 📊 Reconstruir Dashboard | Reconstrói apenas o dashboard (sem re-sincronizar) |
| 🗑️ Limpar tudo e re-sincronizar | Apaga todas as abas e começa do zero |

---

## 🔍 Como filtrar

1. Vai ao separador **`Raw_Kobo`**
2. **Dados → Criar filtro**
3. Filtra pelas colunas:
   - `perfil_participante` — `delegado`, `convidado`, `staff`, `observador`
   - `genero` — `m`, `f`
   - `proveniencia` — nome da igreja
4. Volta ao **Dashboard** → **📊 Kobo Dashboard → Reconstruir Dashboard**

---

## 📁 Estrutura do repositório

```
kobo-assembleia-dashboard/
├── KoboDashboard.gs       # Script principal — colar no Apps Script
├── .env.example           # Exemplo de variáveis de configuração
├── README.md              # Este ficheiro
└── LICENSE
```

---

## 🔧 Campos do formulário suportados

O script suporta automaticamente os campos gerados pelo XLSForm da 27ª Assembleia:

| Campo KoboToolbox | Descrição |
|---|---|
| `perfil_participante` | Delegado / Convidado / Staff / Observador |
| `genero` | m / f |
| `proveniencia` | Igreja de origem |
| `aval_credenciamento` | Avaliação do credenciamento (1–5) |
| `aval_local` | Condições do local (1–5) |
| `aval_alimentacao` | Qualidade da alimentação (1–5) |
| `aval_pontualidade` | Cumprimento do horário (1–5) |
| `aval_temas` | Relevância dos temas (1–5) |
| `aval_metodologia` | Metodologia adotada (1–5) |
| `espaco_opiniao` | sim / não |
| `pontos_fortes` | Texto livre |
| `pontos_melhorar` | Texto livre |

---

## 🐛 Resolução de problemas

| Problema | Causa | Solução |
|---|---|---|
| Dashboard sem dados | Campos com prefixo de grupo não reconhecidos | O script normaliza automaticamente. Abre o editor → Ver → Registos para ver `CAMPOS DISPONÍVEIS` |
| Erro 401 | Token inválido ou expirado | Gera novo token no KoboToolbox → Account → Security |
| Erro 404 | Asset UID errado | Verifica o UID em KoboToolbox → URL do formulário |
| Permissões negadas | Apps Script bloqueado | Vai a [myaccount.google.com/permissions](https://myaccount.google.com/permissions) e autoriza |

---

## 🔐 Segurança das credenciais (produção)

Para evitar expor o token, usa as Properties do Apps Script:

```javascript
// Guardar (correr uma vez manualmente)
function setCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    KOBO_TOKEN    : 'SEU_TOKEN',
    KOBO_ASSET_UID: 'SEU_ASSET_UID',
  });
}

// Usar no script
const CONFIG = {
  TOKEN    : PropertiesService.getScriptProperties().getProperty('KOBO_TOKEN'),
  ASSET_UID: PropertiesService.getScriptProperties().getProperty('KOBO_ASSET_UID'),
  ...
};
```

---

## 📄 Licença

MIT — livre para usar e adaptar.
