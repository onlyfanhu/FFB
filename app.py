"""Curadoria & Votação de Seletivas de Bandas

===========================================
App em Streamlit para curadoria de bandas com gravação no Google Sheets.
"""

import html
import json
import os
import re
import time

from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import requests
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIGURAÇÃO
# ----------------------------------------------------------------------------
COLUMN_MAP = {
    "nome_da_banda": "nome_da_banda",
    "foto_promo": "foto_promo",
    "instagram_da_banda": "instagram_da_banda",
    "streaming_da_banda": "streaming_da_banda",
    "video_da_banda": "video_da_banda",
    "resposta_lineup": "Por que você merece estar no lineup?",
}

CURATOR_SHEET_HEADERS = ["Banda", "Nota Resposta", "Nota Vídeo", "Nota Final"]
AVERAGE_FUNCTION_NAME = "MÉDIA"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(
    page_title="Curadoria de Bandas",
    page_icon="🎸",
    layout="wide",
)

# ----------------------------------------------------------------------------
# CSS - visual limpo e moderno
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .band-name {
        font-size: 2.4rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
        line-height: 1.1;
    }
    .link-pill {
        display: inline-block;
        padding: 0.35rem 0.9rem;
        margin-right: 0.5rem;
        margin-top: 0.4rem;
        border-radius: 999px;
        background: #1f1f2e;
        color: #ffffff !important;
        text-decoration: none !important;
        font-size: 0.85rem;
        font-weight: 600;
        border: 1px solid #3a3a4d;
    }
    .link-pill:hover {
        background: #33334d;
    }
    .progress-tag {
        font-size: 0.85rem;
        color: #888;
        margin-bottom: 0.5rem;
    }
    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        margin-top: 1.2rem;
        margin-bottom: 0.4rem;
    }
    .lineup-box {
        background: #14141f;
        border-left: 4px solid #4f8cff;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-top: 0.6rem;
        font-size: 0.98rem;
        line-height: 1.55;
        color: #e8e8f0;
        white-space: pre-wrap;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# UTILITÁRIOS DE MÍDIA
# ----------------------------------------------------------------------------
def _extract_drive_file_id(url: str) -> str | None:
  if not url:
    return None
  patterns = [
      r"/file/d/([a-zA-Z0-9_-]+)",
      r"[?&]id=([a-zA-Z0-9_-]+)",
      r"/d/([a-zA-Z0-9_-]+)",
  ]
  for pattern in patterns:
    match = re.search(pattern, url)
    if match:
      return match.group(1)
  return None


def _extract_youtube_id(url: str) -> str | None:
  patterns = [
      r"youtu\.be/([a-zA-Z0-9_-]+)",
      r"youtube\.com/watch\?v=([a-zA-Z0-9_-]+)",
      r"youtube\.com/shorts/([a-zA-Z0-9_-]+)",
      r"youtube\.com/embed/([a-zA-Z0-9_-]+)",
  ]
  for pattern in patterns:
    match = re.search(pattern, url)
    if match:
      return match.group(1)
  return None


def _permission_error_message() -> str:
  return (
      "O arquivo não está compartilhado publicamente no Google Drive. "
      "Compartilhe-o como 'Qualquer pessoa com o link' → Leitor."
  )


def convert_drive_image_url(url: str, timeout: int = 10) -> bytes | str | None:
  if not url:
    return None

  file_id = _extract_drive_file_id(url)
  if not file_id:
    return url

  direct_url = f"https://drive.google.com/uc?export=view&id={file_id}"

  try:
    response = requests.get(direct_url, timeout=timeout, allow_redirects=True)

    if "accounts.google.com" in response.url:
      st.session_state["_last_image_error"] = _permission_error_message()
      return None

    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "")

    if "text/html" in content_type:
      confirm_code = None
      if "confirm=" in response.text:
        confirm_code = (
            response.text.split("confirm=")[1].split("&")[0].split('"')[0]
        )

      confirm_url = (
          f"https://drive.google.com/uc?export=download&confirm={confirm_code}&id={file_id}"
          if confirm_code
          else f"https://drive.google.com/uc?export=download&id={file_id}"
      )

      response = requests.get(
          confirm_url, timeout=timeout, allow_redirects=True
      )

      if "accounts.google.com" in response.url:
        st.session_state["_last_image_error"] = _permission_error_message()
        return None

      response.raise_for_status()
      content_type = response.headers.get("Content-Type", "")

    if not content_type.startswith("image/"):
      st.session_state["_last_image_error"] = (
          f"O Drive retornou um conteúdo inesperado ({content_type or 'desconhecido'}). "
          "Verifique se o arquivo é uma imagem pública."
      )
      return None

    st.session_state.pop("_last_image_error", None)
    return response.content

  except requests.RequestException as exc:
    st.session_state["_last_image_error"] = f"Erro de rede ao baixar a imagem: {exc}"
    return None


def convert_media_to_embed(url: str) -> dict:
  if not url:
    return {"type": "empty", "url": ""}

  if "youtube.com" in url or "youtu.be" in url:
    video_id = _extract_youtube_id(url)
    if video_id:
      return {
          "type": "youtube",
          "url": f"https://www.youtube.com/watch?v={video_id}",
      }

  if "drive.google.com" in url:
    file_id = _extract_drive_file_id(url)
    if file_id:
      return {
          "type": "drive_iframe",
          "url": f"https://drive.google.com/file/d/{file_id}/preview",
      }

  return {"type": "direct", "url": url}


# ----------------------------------------------------------------------------
# CONEXÃO COM GOOGLE SHEETS (CORRIGIDO)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_gspread_client():
  creds_data = None
  if "GOOGLE_CREDENTIALS" in os.environ:
    creds_data = json.loads(os.environ["GOOGLE_CREDENTIALS"])
  elif "GOOGLE_CREDENTIALS" in st.secrets:
    creds_data = dict(st.secrets["GOOGLE_CREDENTIALS"])
  else:
    st.error("Variável GOOGLE_CREDENTIALS não encontrada.")
    st.stop()

  credentials = Credentials.from_service_account_info(
      creds_data, scopes=SCOPES
  )
  return gspread.authorize(credentials)


def get_spreadsheet():
  client = get_gspread_client()

  # Lê a URL da planilha configurada no Render (Environment Variables)
  sheet_url = os.environ.get("SHEET_URL")

  if sheet_url:
    return client.open_by_url(sheet_url)

  # Caso não encontre SHEET_URL, abre a primeira planilha associada à conta de serviço
  return client.openall()[0]


def get_main_worksheet() -> gspread.Worksheet:
  """Abre a aba principal (respostas do Forms)."""
  spreadsheet = get_spreadsheet()

  # Lê o nome da aba da variável de ambiente no Render (se existir)
  worksheet_name = os.environ.get("WORKSHEET_NAME")

  if worksheet_name:
    try:
      return spreadsheet.worksheet(worksheet_name)
    except Exception:
      pass

  # Se não houver nome especificado, abre a primeira aba (sheet1)
  return spreadsheet.sheet1


def load_bands_data(force_refresh: bool = False) -> pd.DataFrame:
  if force_refresh:
    st.cache_data.clear()
  return _load_bands_data_cached()


@st.cache_data(ttl=30, show_spinner="Carregando dados das bandas...")
def _load_bands_data_cached() -> pd.DataFrame:
  worksheet = get_main_worksheet()
  records = worksheet.get_all_records()
  return pd.DataFrame(records)


# ----------------------------------------------------------------------------
# ABA INDIVIDUAL DE CADA CURADOR
# ----------------------------------------------------------------------------
def get_or_create_curator_worksheet(curator_name: str) -> gspread.Worksheet:
  spreadsheet = get_spreadsheet()
  safe_name = curator_name.strip()
  try:
    worksheet = spreadsheet.worksheet(safe_name)
  except gspread.exceptions.WorksheetNotFound:
    worksheet = spreadsheet.add_worksheet(
        title=safe_name, rows=200, cols=len(CURATOR_SHEET_HEADERS)
    )
    worksheet.update("A1", [CURATOR_SHEET_HEADERS])
  return worksheet


def load_curator_votes(curator_name: str, force_refresh: bool = False) -> dict:
  if force_refresh:
    st.cache_data.clear()
  return _load_curator_votes_cached(curator_name)


@st.cache_data(ttl=15, show_spinner=False)
def _load_curator_votes_cached(curator_name: str) -> dict:
  worksheet = get_or_create_curator_worksheet(curator_name)
  records = worksheet.get_all_records(expected_headers=CURATOR_SHEET_HEADERS)

  votes = {}
  for i, record in enumerate(records, start=2):
    banda = str(record.get("Banda", "")).strip()
    nota_r = str(record.get("Nota Resposta", "")).strip()
    nota_v = str(record.get("Nota Vídeo", "")).strip()
    if banda and nota_r != "" and nota_v != "":
      votes[banda] = i
  return votes


def save_vote_in_curator_sheet(
    curator_name: str,
    banda_nome: str,
    nota_resposta: float,
    nota_video: float,
) -> None:
  worksheet = get_or_create_curator_worksheet(curator_name)

  existing_row = None
  try:
    cell = worksheet.find(banda_nome, in_column=1)
    existing_row = cell.row if cell else None
  except gspread.exceptions.CellNotFound:
    existing_row = None

  row_num = existing_row if existing_row else len(worksheet.col_values(1)) + 1
  formula_media = f"={AVERAGE_FUNCTION_NAME}(B{row_num}:C{row_num})"

  worksheet.update(
      f"A{row_num}:D{row_num}",
      [[banda_nome, nota_resposta, nota_video, formula_media]],
      value_input_option="USER_ENTERED",
  )


# ----------------------------------------------------------------------------
# INTERFACE STREAMLIT
# ----------------------------------------------------------------------------
def render_header(row: pd.Series) -> None:
  col_img, col_info = st.columns([1, 2.5], gap="large")

  with col_img:
    foto_original = row.get(COLUMN_MAP["foto_promo"], "")
    foto_conteudo = convert_drive_image_url(foto_original)
    if foto_conteudo:
      st.image(foto_conteudo, use_container_width=True)
    elif foto_original:
      erro = st.session_state.get(
          "_last_image_error", "Não foi possível carregar a imagem."
      )
      st.warning(f"⚠️ {erro}")
    else:
      st.info("Sem foto promocional cadastrada.")

  with col_info:
    st.markdown(
        f'<div class="band-name">{row.get(COLUMN_MAP["nome_da_banda"], "Banda sem nome")}</div>',
        unsafe_allow_html=True,
    )

    instagram = row.get(COLUMN_MAP["instagram_da_banda"], "")
    streaming = row.get(COLUMN_MAP["streaming_da_banda"], "")

    links_html = ""
    if instagram:
      links_html += (
          f'<a class="link-pill" href="{instagram}" target="_blank">📸'
          ' Instagram</a>'
      )
    if streaming:
      links_html += (
          f'<a class="link-pill" href="{streaming}" target="_blank">🎧'
          ' Streaming</a>'
      )
    if links_html:
      st.markdown(links_html, unsafe_allow_html=True)


def render_video(row: pd.Series) -> None:
  st.markdown(
      '<div class="section-title">🎬 Vídeo da banda</div>',
      unsafe_allow_html=True,
  )
  video_url = row.get(COLUMN_MAP["video_da_banda"], "")
  media = convert_media_to_embed(video_url)

  if media["type"] in ("youtube", "direct"):
    st.video(media["url"])
  elif media["type"] == "drive_iframe":
    st.components.v1.iframe(media["url"], height=480)
    st.caption(
        "Se o vídeo não carregar, verifique o compartilhamento no Drive. "
        f"[Abrir vídeo diretamente no Drive]({video_url})"
    )
  else:
    st.warning("Nenhum vídeo cadastrado para esta banda.")


def render_voting_form(
    curator_name: str,
    banda_nome: str,
    row: pd.Series,
    pending_bandas: list[str],
) -> None:
  st.markdown(
      '<div class="section-title">✅ Avaliação</div>', unsafe_allow_html=True
  )

  resposta_lineup = row.get(COLUMN_MAP["resposta_lineup"], "")
  st.markdown("**📝 Por que a banda acha que merece estar no lineup:**")
  if str(resposta_lineup).strip():
    texto_seguro = html.escape(str(resposta_lineup))
    st.markdown(
        f'<div class="lineup-box">{texto_seguro}</div>', unsafe_allow_html=True
    )
  else:
    st.info("Nenhuma resposta cadastrada para esta pergunta.")

  with st.form(key=f"vote_form_{banda_nome}"):
    c1, c2 = st.columns(2)
    with c1:
      nota_resposta = st.number_input(
          "Nota da Redação/Resposta",
          min_value=0.0,
          max_value=10.0,
          value=5.0,
          step=0.5,
      )
    with c2:
      nota_video = st.number_input(
          "Nota do Vídeo",
          min_value=0.0,
          max_value=10.0,
          value=5.0,
          step=0.5,
      )

    submitted = st.form_submit_button(
        "➡️ Enviar Notas e Próxima Banda", use_container_width=True
    )

  if submitted:
    with st.spinner(f"Salvando notas na aba de {curator_name}..."):
      save_vote_in_curator_sheet(
          curator_name, banda_nome, nota_resposta, nota_video
      )
      time.sleep(0.3)
      load_curator_votes(curator_name, force_refresh=True)

    st.success("Notas registradas com sucesso!")
    remaining = [b for b in pending_bandas if b != banda_nome]
    st.session_state["current_banda"] = remaining[0] if remaining else None
    st.rerun()


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main() -> None:
  st.title("🎸 Curadoria de Seletivas de Bandas")

  with st.sidebar:
    st.header("⚙️ Painel")
    curator_name = st.text_input(
        "Seu nome (curador)",
        value=st.session_state.get("curator_name", ""),
        placeholder="Ex: Caio",
    ).strip()
    st.session_state["curator_name"] = curator_name

  if not curator_name:
    st.info(
        "👋 Digite seu nome na barra lateral para começar a avaliar as bandas."
    )
    return

  try:
    get_or_create_curator_worksheet(curator_name)
  except Exception as exc:
    st.error(
        f"Não foi possível criar/abrir a aba de '{curator_name}' na planilha."
    )
    st.exception(exc)
    return

  try:
    df = load_bands_data()
  except Exception as exc:
    st.error("Não foi possível carregar a planilha principal.")
    st.exception(exc)
    return

  if df.empty:
    st.warning("A planilha principal ainda não tem nenhuma resposta.")
    return

  curator_votes = load_curator_votes(curator_name)
  todas_bandas = df[COLUMN_MAP["nome_da_banda"]].astype(str).tolist()
  pending_bandas = [b for b in todas_bandas if b not in curator_votes]

  if (
      "current_banda" not in st.session_state
      or st.session_state["current_banda"] not in pending_bandas
  ):
    st.session_state["current_banda"] = (
        pending_bandas[0] if pending_bandas else None
    )

  current_banda = st.session_state["current_banda"]
  total = len(todas_bandas)
  votadas = total - len(pending_bandas)

  with st.sidebar:
    st.metric("Total de bandas", total)
    st.metric(f"Avaliadas por {curator_name}", votadas)
    st.metric("Pendentes", len(pending_bandas))
    if st.button("🔄 Recarregar dados"):
      load_bands_data(force_refresh=True)
      load_curator_votes(curator_name, force_refresh=True)
      st.rerun()

  if current_banda is None:
    st.balloons()
    st.success(f"🎉 {curator_name}, você já avaliou todas as {total} bandas!")
    return

  st.markdown(
      f'<div class="progress-tag">Banda pendente {votadas + 1} de {total}'
      f" &nbsp;·&nbsp; {len(pending_bandas)} restantes &nbsp;·&nbsp; curador:"
      f" <b>{curator_name}</b></div>",
      unsafe_allow_html=True,
  )
  st.progress(votadas / total if total else 0)

  row = df[df[COLUMN_MAP["nome_da_banda"]].astype(str) == current_banda].iloc[0]

  render_header(row)
  st.divider()
  render_video(row)
  st.divider()
  render_voting_form(curator_name, current_banda, row, pending_bandas)


if __name__ == "__main__":
  main()
