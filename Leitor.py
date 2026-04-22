from __future__ import annotations

import argparse
import configparser
import ctypes
import queue
import re
import shlex
import shutil
import subprocess
import sys
import threading
import time
import unicodedata
from contextlib import contextmanager
from ctypes import wintypes
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Sequence

import fdb
import psutil
import tkinter as tk
from tkinter import font as tkfont, messagebox, scrolledtext, ttk


# Evita conflito entre o dialogo nativo do Windows e a inicializacao padrao do pywinauto.
sys.coinit_flags = 2

USER32_BLOQUEIO_ENTRADA = ctypes.WinDLL("user32", use_last_error=True)
USER32_BLOQUEIO_ENTRADA.BlockInput.argtypes = [wintypes.BOOL]
USER32_BLOQUEIO_ENTRADA.BlockInput.restype = wintypes.BOOL

try:
    import win32api
    import win32con
    import win32gui
    import win32process
except ImportError:
    USER32 = ctypes.WinDLL("user32", use_last_error=True)
    KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    USER32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    USER32.GetWindowThreadProcessId.restype = wintypes.DWORD
    USER32.IsWindowVisible.argtypes = [wintypes.HWND]
    USER32.IsWindowVisible.restype = wintypes.BOOL
    USER32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    USER32.GetWindowTextLengthW.restype = ctypes.c_int
    USER32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    USER32.GetWindowTextW.restype = ctypes.c_int
    USER32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    USER32.GetClassNameW.restype = ctypes.c_int
    USER32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
    USER32.EnumWindows.restype = wintypes.BOOL
    USER32.EnumChildWindows.argtypes = [wintypes.HWND, WNDENUMPROC, wintypes.LPARAM]
    USER32.EnumChildWindows.restype = wintypes.BOOL
    USER32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    USER32.ShowWindow.restype = wintypes.BOOL
    USER32.SetForegroundWindow.argtypes = [wintypes.HWND]
    USER32.SetForegroundWindow.restype = wintypes.BOOL
    USER32.IsWindow.argtypes = [wintypes.HWND]
    USER32.IsWindow.restype = wintypes.BOOL
    USER32.SendMessageW.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
    USER32.SendMessageW.restype = wintypes.LPARAM
    USER32.PostMessageW.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
    USER32.PostMessageW.restype = wintypes.BOOL
    USER32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    USER32.GetWindowRect.restype = wintypes.BOOL
    USER32.GetClientRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    USER32.GetClientRect.restype = wintypes.BOOL
    USER32.ClientToScreen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
    USER32.ClientToScreen.restype = wintypes.BOOL
    USER32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
    USER32.SetCursorPos.restype = wintypes.BOOL
    USER32.mouse_event.argtypes = [
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.DWORD,
        ctypes.c_ulong,
    ]
    USER32.mouse_event.restype = None
    KERNEL32.GetLogicalDriveStringsW.argtypes = [wintypes.DWORD, wintypes.LPWSTR]
    KERNEL32.GetLogicalDriveStringsW.restype = wintypes.DWORD

    class _Win32ConCompat:
        SW_RESTORE = 9
        BM_GETCHECK = 0x00F0
        BST_CHECKED = 1
        WM_CLOSE = 0x0010
        WM_MOUSEMOVE = 0x0200
        WM_LBUTTONDOWN = 0x0201
        WM_LBUTTONUP = 0x0202
        MK_LBUTTON = 0x0001
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP = 0x0004

    class _Win32ProcessCompat:
        @staticmethod
        def GetWindowThreadProcessId(hwnd: int) -> tuple[int, int]:
            pid = wintypes.DWORD()
            thread_id = USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(thread_id), int(pid.value)

    class _Win32GuiCompat:
        @staticmethod
        def IsWindowVisible(hwnd: int) -> bool:
            return bool(USER32.IsWindowVisible(hwnd))

        @staticmethod
        def GetWindowText(hwnd: int) -> str:
            tamanho = max(1, USER32.GetWindowTextLengthW(hwnd) + 1)
            buffer = ctypes.create_unicode_buffer(tamanho)
            USER32.GetWindowTextW(hwnd, buffer, tamanho)
            return buffer.value

        @staticmethod
        def GetClassName(hwnd: int) -> str:
            buffer = ctypes.create_unicode_buffer(256)
            USER32.GetClassNameW(hwnd, buffer, len(buffer))
            return buffer.value

        @staticmethod
        def EnumWindows(callback: Callable[[int, object], bool], extra: object) -> None:
            def callback_ctypes(hwnd: int, lparam: int) -> bool:
                return bool(callback(int(hwnd), extra))

            proc = WNDENUMPROC(callback_ctypes)
            USER32.EnumWindows(proc, 0)

        @staticmethod
        def EnumChildWindows(hwnd_janela: int, callback: Callable[[int, object], bool], extra: object) -> None:
            def callback_ctypes(hwnd: int, lparam: int) -> bool:
                return bool(callback(int(hwnd), extra))

            proc = WNDENUMPROC(callback_ctypes)
            USER32.EnumChildWindows(hwnd_janela, proc, 0)

        @staticmethod
        def ShowWindow(hwnd: int, comando: int) -> bool:
            return bool(USER32.ShowWindow(hwnd, comando))

        @staticmethod
        def SetForegroundWindow(hwnd: int) -> bool:
            return bool(USER32.SetForegroundWindow(hwnd))

        @staticmethod
        def IsWindow(hwnd: int) -> bool:
            return bool(USER32.IsWindow(hwnd))

        @staticmethod
        def SendMessage(hwnd: int, mensagem: int, wparam: int, lparam: int) -> int:
            return int(USER32.SendMessageW(hwnd, mensagem, wparam, lparam))

        @staticmethod
        def PostMessage(hwnd: int, mensagem: int, wparam: int, lparam: int) -> bool:
            return bool(USER32.PostMessageW(hwnd, mensagem, wparam, lparam))

        @staticmethod
        def GetWindowRect(hwnd: int) -> tuple[int, int, int, int]:
            rect = wintypes.RECT()
            if not USER32.GetWindowRect(hwnd, ctypes.byref(rect)):
                raise ctypes.WinError(ctypes.get_last_error())
            return rect.left, rect.top, rect.right, rect.bottom

        @staticmethod
        def GetClientRect(hwnd: int) -> tuple[int, int, int, int]:
            rect = wintypes.RECT()
            if not USER32.GetClientRect(hwnd, ctypes.byref(rect)):
                raise ctypes.WinError(ctypes.get_last_error())
            return rect.left, rect.top, rect.right, rect.bottom

        @staticmethod
        def ClientToScreen(hwnd: int, ponto: tuple[int, int]) -> tuple[int, int]:
            point = wintypes.POINT(*ponto)
            if not USER32.ClientToScreen(hwnd, ctypes.byref(point)):
                raise ctypes.WinError(ctypes.get_last_error())
            return point.x, point.y

    class _Win32ApiCompat:
        @staticmethod
        def MAKELONG(low: int, high: int) -> int:
            return ((int(high) & 0xFFFF) << 16) | (int(low) & 0xFFFF)

        @staticmethod
        def SetCursorPos(ponto: tuple[int, int]) -> None:
            if not USER32.SetCursorPos(int(ponto[0]), int(ponto[1])):
                raise ctypes.WinError(ctypes.get_last_error())

        @staticmethod
        def mouse_event(evento: int, dx: int, dy: int, data: int, extra: int) -> None:
            USER32.mouse_event(evento, dx, dy, data, extra)

        @staticmethod
        def GetLogicalDriveStrings() -> str:
            tamanho = KERNEL32.GetLogicalDriveStringsW(0, None)
            buffer = ctypes.create_unicode_buffer(tamanho)
            KERNEL32.GetLogicalDriveStringsW(tamanho, buffer)
            return ctypes.wstring_at(buffer, tamanho)

    win32api = _Win32ApiCompat()
    win32con = _Win32ConCompat()
    win32gui = _Win32GuiCompat()
    win32process = _Win32ProcessCompat()

PASTA_PADRAO = Path(r"C:\interface")
EXTENSOES_FIREBIRD = {".fdb", ".gdb", ".ib", ".fb", ".fbd"}
CAMPO_COMPILACAO = "T000_TS_COMPILACAO"
NOME_BASE_ISOLADA = "BASE EM ATUALIZACAO.FDB"
NOME_BASE_ORIGINAL = "BD_INTERFACE.FDB"
NOME_ARQUIVO_MAPEAMENTO_BASE = "mapeamento_base_em_atualizacao.txt"
INDICADORES_CAMINHO_BANCO_IGNORADO = ("interface backup", "interfacebackup")
CHAVE_FILA_VAZIA = "Nenhuma atualização pendente."
APP_TITULO = "Interface Update"
NOME_PASTA_SETUPS_PREFERENCIAL = "SETUP INTERFACE ATUALIZACAO"
NOMES_PASTA_SETUPS_NORMALIZADOS = {"setup interface atualizacao"}
ENTERS_POR_SETUP = 5
TEMPO_ESPERA_ENTRE_ENTERS = 1.2
TEMPO_LIMITE_JANELA_INSTALADOR = 120
TEMPO_LIMITE_CONCLUSAO_INSTALADOR = 900
TEMPO_LIMITE_ABERTURA_INTERFACE = 300
TEMPO_ESPERA_FECHAMENTO_INTERFACE = 3
TEMPO_ESPERA_REACAO_FECHAMENTO = 1.0
TEMPO_LIMITE_LOCALIZAR_CANCELAR = 20


def pasta_recursos() -> Path:
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent


def pasta_aplicacao() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalizar_nome_arquivo(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    return re.sub(r"\s+", " ", texto.casefold()).strip()


def primeira_fonte_disponivel(candidatas: Sequence[str], padrao: str) -> str:
    try:
        familias = {nome.casefold() for nome in tkfont.families()}
    except tk.TclError:
        familias = set()

    for candidata in candidatas:
        if candidata.casefold() in familias:
            return candidata
    return padrao


def listar_bases_busca_setups(base_dir: Path, niveis_acima: int = 2) -> list[Path]:
    bases: list[Path] = []
    vistos: set[str] = set()
    atual = base_dir

    for _ in range(max(0, niveis_acima) + 1):
        try:
            existe = atual.exists() and atual.is_dir()
        except OSError:
            existe = False
        if existe:
            chave = str(atual.resolve()).lower()
            if chave not in vistos:
                vistos.add(chave)
                bases.append(atual)

        proximo = atual.parent
        if proximo == atual:
            break
        atual = proximo

    return bases


def encontrar_pasta_setups_semelhante(base_dir: Path, nome_desejado: str | None = None) -> Path | None:
    alvo_normalizado = normalizar_nome_arquivo(nome_desejado or NOME_PASTA_SETUPS_PREFERENCIAL)

    try:
        diretorios = [item for item in base_dir.iterdir() if item.is_dir()]
    except OSError:
        return None

    candidatos_exatos: list[Path] = []
    candidatos_parciais: list[Path] = []

    for diretorio in diretorios:
        nome_normalizado = normalizar_nome_arquivo(diretorio.name)
        if nome_normalizado in NOMES_PASTA_SETUPS_NORMALIZADOS or nome_normalizado == alvo_normalizado:
            candidatos_exatos.append(diretorio)
            continue
        if "setup interface" in nome_normalizado and "atualizacao" in nome_normalizado:
            candidatos_parciais.append(diretorio)

    if candidatos_exatos:
        return sorted(candidatos_exatos, key=lambda item: item.name)[0]
    if candidatos_parciais:
        return sorted(candidatos_parciais, key=lambda item: item.name)[0]
    return None


def detectar_pasta_setups_padrao(base_dir: Path) -> Path:
    for base_busca in listar_bases_busca_setups(base_dir):
        encontrada = encontrar_pasta_setups_semelhante(base_busca)
        if encontrada is not None:
            return encontrada
    return base_dir / NOME_PASTA_SETUPS_PREFERENCIAL


PASTA_RECURSOS = pasta_recursos()
PASTA_APLICACAO = pasta_aplicacao()
PASTA_SETUPS_PADRAO = detectar_pasta_setups_padrao(PASTA_APLICACAO)
ARQUIVO_ICONE = PASTA_RECURSOS / "assets" / "app_icon.ico"
ARQUIVO_LOGO = PASTA_RECURSOS / "assets" / "app_logo.png"
ARQUIVO_SQL_CORRECAO_GRID = PASTA_RECURSOS / "assets" / "sql" / "correcao_grid_localizacao_produtos.sql"


@dataclass(frozen=True)
class BancoCandidato:
    caminho: Path
    usuario: str = "SYSDBA"
    senha: str = "masterkey"
    origem: str = ""
    referenciado_em_config: bool = False


@dataclass(frozen=True)
class VersaoBanco:
    valor_bruto: object
    compilado_em: datetime
    tem_hora: bool


@dataclass(frozen=True)
class SetupAtualizacao:
    caminho: Path
    compilado_em: datetime

    @property
    def descricao(self) -> str:
        return f"{self.compilado_em.strftime('%d/%m/%Y %H:%M')} - {self.caminho.name}"


@dataclass(frozen=True)
class AlteracaoArquivo:
    caminho: Path
    conteudo_original: str
    conteudo_novo: str
    ocorrencias: int


@dataclass(frozen=True)
class ResultadoAnalise:
    pasta_interface: Path
    pasta_setups: Path
    banco_principal: BancoCandidato
    tabela_compilacao: str
    versao_atual: VersaoBanco
    fila_atualizacao: list[SetupAtualizacao]
    setups_disponiveis: list[SetupAtualizacao]
    base_isolada: bool


@dataclass(frozen=True)
class JanelaDetectada:
    handle: int
    pid: int
    titulo: str
    classe: str


@dataclass(frozen=True)
class ControleDetectado:
    handle: int
    classe: str
    texto: str
    visivel: bool
    retangulo_tela: tuple[int, int, int, int] | None = None


class OperacaoCancelada(RuntimeError):
    pass


def logar(logger: Callable[[str], None] | None, mensagem: str) -> None:
    if logger is not None:
        logger(mensagem)


def verificar_cancelamento(cancelar_evento: threading.Event | None, contexto: str = "") -> None:
    if cancelar_evento is not None and cancelar_evento.is_set():
        if contexto:
            raise OperacaoCancelada(f"Operacao interrompida durante {contexto}.")
        raise OperacaoCancelada("Operacao interrompida pelo usuario.")


def dormir_interrompivel(
    segundos: float,
    cancelar_evento: threading.Event | None = None,
    contexto: str = "",
) -> None:
    fim = time.monotonic() + max(0.0, segundos)
    while True:
        verificar_cancelamento(cancelar_evento, contexto)
        restante = fim - time.monotonic()
        if restante <= 0:
            return
        time.sleep(min(0.2, restante))


def normalizar_texto_ui(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    return texto.casefold().strip()


def usuario_e_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:  # noqa: BLE001
        return False


def acao_requer_elevacao(acao: str) -> bool:
    return acao in {"gui", "preparar", "restaurar", "atualizar"}


def alterar_bloqueio_entrada_usuario(bloquear: bool) -> None:
    ctypes.set_last_error(0)
    if USER32_BLOQUEIO_ENTRADA.BlockInput(bool(bloquear)):
        return

    codigo = ctypes.get_last_error()
    acao = "bloquear" if bloquear else "liberar"
    detalhe = ctypes.FormatError(codigo).strip() if codigo else "retorno inesperado do Windows"
    raise RuntimeError(f"Nao foi possivel {acao} teclado e mouse: {detalhe}.")


@contextmanager
def bloquear_entrada_usuario(
    logger: Callable[[str], None] | None = None,
    contexto: str = "a automacao",
):
    logar(logger, f"Bloqueando teclado e mouse durante {contexto}.")
    alterar_bloqueio_entrada_usuario(True)
    try:
        yield
    finally:
        try:
            alterar_bloqueio_entrada_usuario(False)
        except Exception as exc:  # noqa: BLE001
            logar(logger, f"Aviso: nao foi possivel liberar teclado e mouse automaticamente: {exc}")
        else:
            logar(logger, "Teclado e mouse liberados.")


def relancar_como_admin() -> bool:
    if getattr(sys, "frozen", False):
        executavel = sys.executable
        argumentos = sys.argv[1:]
    else:
        executavel = sys.executable
        argumentos = [str(Path(__file__).resolve()), *sys.argv[1:]]

    parametros = subprocess.list2cmdline(argumentos)
    resultado = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        executavel,
        parametros,
        str(PASTA_APLICACAO),
        1,
    )
    return resultado > 32


def coletar_pids_relacionados(pids_iniciais: Iterable[int]) -> list[int]:
    conhecidos = {int(pid) for pid in pids_iniciais if pid}

    for _ in range(4):
        atualizados = set(conhecidos)
        for pid in list(conhecidos):
            if not psutil.pid_exists(pid):
                continue

            try:
                processo = psutil.Process(pid)
                atualizados.update(filho.pid for filho in processo.children(recursive=True))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if atualizados == conhecidos:
            break
        conhecidos = atualizados

    return sorted(pid for pid in conhecidos if psutil.pid_exists(pid))


def listar_janelas_visiveis(pids: Iterable[int]) -> list[JanelaDetectada]:
    pids_validos = set(int(pid) for pid in pids if pid)
    janelas: list[JanelaDetectada] = []

    def callback(hwnd: int, _extra: object) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in pids_validos or not win32gui.IsWindowVisible(hwnd):
                return True

            janelas.append(
                JanelaDetectada(
                    handle=hwnd,
                    pid=pid,
                    titulo=win32gui.GetWindowText(hwnd),
                    classe=win32gui.GetClassName(hwnd),
                )
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(callback, None)
    return janelas


def listar_todas_janelas_visiveis() -> list[JanelaDetectada]:
    janelas: list[JanelaDetectada] = []

    def callback(hwnd: int, _extra: object) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if not win32gui.IsWindowVisible(hwnd):
                return True

            janelas.append(
                JanelaDetectada(
                    handle=hwnd,
                    pid=pid,
                    titulo=win32gui.GetWindowText(hwnd),
                    classe=win32gui.GetClassName(hwnd),
                )
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(callback, None)
    return janelas


def listar_controles(hwnd_janela: int) -> list[ControleDetectado]:
    controles: list[ControleDetectado] = []

    def callback(hwnd: int, _extra: object) -> bool:
        try:
            controles.append(
                ControleDetectado(
                    handle=hwnd,
                    classe=win32gui.GetClassName(hwnd),
                    texto=win32gui.GetWindowText(hwnd),
                    visivel=bool(win32gui.IsWindowVisible(hwnd)),
                    retangulo_tela=obter_retangulo_tela(hwnd),
                )
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumChildWindows(hwnd_janela, callback, None)
    return controles


def aguardar_condicao(
    descricao: str,
    funcao: Callable[[], object],
    timeout: float,
    intervalo: float = 0.4,
    cancelar_evento: threading.Event | None = None,
) -> object:
    inicio = time.monotonic()
    ultimo_erro: Exception | None = None

    while time.monotonic() - inicio <= timeout:
        verificar_cancelamento(cancelar_evento, descricao)
        try:
            resultado = funcao()
        except Exception as exc:  # noqa: BLE001
            ultimo_erro = exc
        else:
            if resultado:
                return resultado
        dormir_interrompivel(intervalo, cancelar_evento, descricao)

    if ultimo_erro is not None:
        raise RuntimeError(f"Tempo esgotado ao aguardar {descricao}: {ultimo_erro}") from ultimo_erro
    raise RuntimeError(f"Tempo esgotado ao aguardar {descricao}.")


def aguardar_janela(
    origem_pids: Iterable[int] | Callable[[], Iterable[int]],
    descricao: str,
    timeout: float,
    predicado: Callable[[JanelaDetectada], bool],
    cancelar_evento: threading.Event | None = None,
) -> JanelaDetectada:
    def buscar() -> JanelaDetectada | None:
        pids_base = origem_pids() if callable(origem_pids) else origem_pids
        pids = coletar_pids_relacionados(pids_base)
        for janela in listar_janelas_visiveis(pids):
            if predicado(janela):
                return janela
        return None

    resultado = aguardar_condicao(descricao, buscar, timeout, cancelar_evento=cancelar_evento)
    assert isinstance(resultado, JanelaDetectada)
    return resultado


def trazer_janela_para_frente(hwnd: int) -> None:
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:  # noqa: BLE001
        pass

    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:  # noqa: BLE001
        pass


def janela_existe(hwnd: int) -> bool:
    try:
        return bool(win32gui.IsWindow(hwnd))
    except Exception:  # noqa: BLE001
        return False


def janela_continua_visivel(hwnd: int) -> bool:
    if not janela_existe(hwnd):
        return False

    try:
        return bool(win32gui.IsWindowVisible(hwnd))
    except Exception:  # noqa: BLE001
        return False


def obter_bbox_cliente_em_tela(hwnd: int) -> tuple[int, int, int, int]:
    esquerda, topo, direita, base = win32gui.GetClientRect(hwnd)
    canto_superior = win32gui.ClientToScreen(hwnd, (esquerda, topo))
    canto_inferior = win32gui.ClientToScreen(hwnd, (direita, base))
    return (
        int(canto_superior[0]),
        int(canto_superior[1]),
        int(canto_inferior[0]),
        int(canto_inferior[1]),
    )


def obter_retangulo_tela(hwnd: int) -> tuple[int, int, int, int] | None:
    try:
        esquerda, topo, direita, base = win32gui.GetWindowRect(hwnd)
    except Exception:  # noqa: BLE001
        return None
    return int(esquerda), int(topo), int(direita), int(base)


def obter_centro_controle_em_cliente(
    hwnd_janela: int,
    controle: ControleDetectado,
) -> tuple[int, int] | None:
    retangulo = controle.retangulo_tela or obter_retangulo_tela(controle.handle)
    if retangulo is None:
        return None

    origem_cliente_tela = win32gui.ClientToScreen(hwnd_janela, (0, 0))
    esquerda, topo, direita, base = retangulo
    centro_tela = ((esquerda + direita) // 2, (topo + base) // 2)
    return (
        int(centro_tela[0] - origem_cliente_tela[0]),
        int(centro_tela[1] - origem_cliente_tela[1]),
    )


def _faixas_com_minimo(contagens: list[int], minimo: int, largura_minima: int) -> list[tuple[int, int]]:
    faixas: list[tuple[int, int]] = []
    inicio: int | None = None

    for indice, valor in enumerate([*contagens, 0]):
        if valor >= minimo:
            if inicio is None:
                inicio = indice
            continue

        if inicio is not None and indice - inicio >= largura_minima:
            faixas.append((inicio, indice - 1))
        inicio = None

    return faixas


def pixel_parece_botao_azul(pixel: tuple[int, ...]) -> bool:
    vermelho, verde, azul = pixel[:3]
    return azul >= 140 and verde >= 95 and vermelho <= 90 and (azul - vermelho) >= 60 and (verde - vermelho) >= 35


def localizar_botao_cancelar_por_imagem(
    hwnd_janela: int,
    logger: Callable[[str], None] | None = None,
) -> tuple[int, int] | None:
    try:
        from PIL import ImageGrab
    except ImportError:
        logar(logger, "Reconhecimento visual indisponivel: biblioteca Pillow nao encontrada.")
        return None

    try:
        bbox = obter_bbox_cliente_em_tela(hwnd_janela)
        imagem = ImageGrab.grab(bbox=bbox)
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Nao foi possivel capturar a janela para reconhecimento visual: {exc}")
        return None

    largura, altura = imagem.size
    if largura <= 0 or altura <= 0:
        return None

    origem_x = max(0, int(largura * 0.05))
    origem_y = max(0, int(altura * 0.35))
    limite_x = min(largura, int(largura * 0.95))
    limite_y = min(altura, int(altura * 0.95))
    regiao = imagem.crop((origem_x, origem_y, limite_x, limite_y))
    largura_regiao, altura_regiao = regiao.size
    pixels = regiao.load()

    contagem_por_coluna = [
        sum(1 for y in range(altura_regiao) if pixel_parece_botao_azul(pixels[x, y]))
        for x in range(largura_regiao)
    ]
    faixas_x = _faixas_com_minimo(
        contagem_por_coluna,
        minimo=max(6, altura_regiao // 10),
        largura_minima=max(40, largura_regiao // 12),
    )

    candidatos: list[tuple[int, int, int, int, int, int]] = []
    for inicio_x, fim_x in faixas_x:
        contagem_por_linha = [
            sum(1 for x in range(inicio_x, fim_x + 1) if pixel_parece_botao_azul(pixels[x, y]))
            for y in range(altura_regiao)
        ]
        faixas_y = _faixas_com_minimo(
            contagem_por_linha,
            minimo=max(8, (fim_x - inicio_x + 1) // 6),
            largura_minima=max(18, altura_regiao // 10),
        )
        if not faixas_y:
            continue

        inicio_y, fim_y = max(faixas_y, key=lambda faixa: faixa[1] - faixa[0])
        largura_box = fim_x - inicio_x + 1
        altura_box = fim_y - inicio_y + 1
        if largura_box < 60 or altura_box < 24:
            continue
        proporcao = largura_box / max(1, altura_box)
        if proporcao < 1.2 or proporcao > 6.5:
            continue

        area_azul = sum(
            1
            for x in range(inicio_x, fim_x + 1)
            for y in range(inicio_y, fim_y + 1)
            if pixel_parece_botao_azul(pixels[x, y])
        )
        if area_azul < 1200:
            continue

        centro_x = (inicio_x + fim_x) // 2
        centro_y = (inicio_y + fim_y) // 2
        score = area_azul + (centro_x * 2) + centro_y
        candidatos.append((inicio_x, inicio_y, fim_x, fim_y, area_azul, score))

    if not candidatos:
        logar(logger, "Reconhecimento visual nao encontrou um botao azul compatível com 'Cancelar'.")
        return None

    inicio_x, inicio_y, fim_x, fim_y, _area_azul, _score = max(
        candidatos,
        key=lambda item: item[5],
    )
    alvo_cliente = (
        origem_x + ((inicio_x + fim_x) // 2),
        origem_y + ((inicio_y + fim_y) // 2),
    )
    logar(logger, f"Reconhecimento visual encontrou o botao Cancelar na coordenada {alvo_cliente}.")
    return alvo_cliente


def localizar_botao_cancelar_por_geometria(
    hwnd_janela: int,
    logger: Callable[[str], None] | None = None,
) -> tuple[int, int] | None:
    try:
        esquerda, topo, direita, base = win32gui.GetClientRect(hwnd_janela)
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Nao foi possivel ler a geometria da janela para localizar Cancelar: {exc}")
        return None

    largura = max(1, direita - esquerda)
    altura = max(1, base - topo)
    candidatos: list[tuple[int, tuple[int, int], ControleDetectado]] = []

    for controle in listar_controles(hwnd_janela):
        if not controle.visivel:
            continue

        classe = normalizar_texto_ui(controle.classe)
        if "button" not in classe:
            continue

        try:
            centro_cliente = obter_centro_controle_em_cliente(hwnd_janela, controle)
            retangulo = controle.retangulo_tela or obter_retangulo_tela(controle.handle)
        except Exception:  # noqa: BLE001
            continue
        if centro_cliente is None or retangulo is None:
            continue

        centro_x, centro_y = centro_cliente
        if not (0 <= centro_x <= largura and 0 <= centro_y <= altura):
            continue

        esquerda_tela, topo_tela, direita_tela, base_tela = retangulo
        largura_controle = max(1, direita_tela - esquerda_tela)
        altura_controle = max(1, base_tela - topo_tela)
        if largura_controle < 50 or altura_controle < 18:
            continue
        if largura_controle / max(1, altura_controle) > 8:
            continue

        texto = normalizar_texto_ui(controle.texto)
        score = (centro_y * 3) + (centro_x * 2)
        if any(termo in texto for termo in ("cancelar", "cancel", "fechar", "sair")):
            score += largura * altura
        elif texto and any(termo in texto for termo in ("ok", "entrar", "login", "confirmar")):
            score -= largura

        candidatos.append((score, centro_cliente, controle))

    if not candidatos:
        logar(logger, "Nenhum controle compativel com o botao Cancelar foi encontrado pela geometria da janela.")
        return None

    _score, alvo_cliente, controle = max(candidatos, key=lambda item: item[0])
    logar(
        logger,
        "Heuristica geometrica selecionou o controle "
        f"'{controle.texto or controle.classe or controle.handle}' na coordenada {alvo_cliente}.",
    )
    return alvo_cliente


def clicar_na_coordenada_cliente(
    hwnd_janela: int,
    alvo_cliente: tuple[int, int],
    descricao: str,
    logger: Callable[[str], None] | None = None,
) -> None:
    alvo_tela = win32gui.ClientToScreen(hwnd_janela, alvo_cliente)
    lparam = win32api.MAKELONG(alvo_cliente[0], alvo_cliente[1])

    trazer_janela_para_frente(hwnd_janela)
    try:
        win32gui.SendMessage(hwnd_janela, win32con.WM_MOUSEMOVE, 0, lparam)
        win32gui.SendMessage(hwnd_janela, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        win32gui.SendMessage(hwnd_janela, win32con.WM_LBUTTONUP, 0, lparam)
        logar(logger, f"{descricao} acionado por mensagem na coordenada {alvo_cliente}.")
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Falha ao enviar mensagem de clique para {descricao}: {exc}")

    try:
        time.sleep(0.4)
        win32api.SetCursorPos(alvo_tela)
        time.sleep(0.15)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        logar(logger, f"{descricao} acionado por clique fisico na coordenada {alvo_cliente}.")
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Falha ao executar o clique fisico para {descricao}: {exc}")


def enviar_escape_para_janela(hwnd_janela: int, logger: Callable[[str], None] | None = None) -> bool:
    trazer_janela_para_frente(hwnd_janela)

    try:
        wrapper = obter_wrapper_janela(hwnd_janela)
    except Exception:  # noqa: BLE001
        wrapper = None

    try:
        if wrapper is not None:
            wrapper.set_focus()
            wrapper.type_keys("{ESC}", set_foreground=True)
        else:
            enviar_teclas_pywinauto("{ESC}")
        logar(logger, "Tecla ESC enviada para a tela de acesso da Interface.")
        return True
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Nao foi possivel enviar ESC para a tela de acesso: {exc}")
        return False


def _application_pywinauto():
    try:
        from pywinauto.application import Application
    except ImportError as exc:
        raise RuntimeError(
            "A automacao da Interface requer a biblioteca 'pywinauto'. "
            "Instale com: pip install pywinauto"
        ) from exc
    return Application


def enviar_teclas_pywinauto(teclas: str) -> None:
    try:
        from pywinauto.keyboard import send_keys
    except ImportError as exc:
        raise RuntimeError(
            "A automacao da Interface requer a biblioteca 'pywinauto'. "
            "Instale com: pip install pywinauto"
        ) from exc
    send_keys(teclas)


def obter_wrapper_janela(hwnd: int):
    app = _application_pywinauto()(backend="win32").connect(handle=hwnd)
    return app.window(handle=hwnd)


def localizar_controle_por_texto(
    hwnd_janela: int,
    trechos: Sequence[str],
    classes_permitidas: Sequence[str] | None = None,
) -> ControleDetectado | None:
    termos = [normalizar_texto_ui(trecho) for trecho in trechos]
    classes = [normalizar_texto_ui(classe) for classe in (classes_permitidas or ())]

    for controle in listar_controles(hwnd_janela):
        texto = normalizar_texto_ui(controle.texto)
        classe = normalizar_texto_ui(controle.classe)
        if not texto:
            continue
        if classes and not any(classe_permitida in classe for classe_permitida in classes):
            continue
        if any(termo in texto for termo in termos):
            return controle

    return None


def acionar_controle(hwnd_janela: int, controle: ControleDetectado) -> None:
    janela = obter_wrapper_janela(hwnd_janela)
    elemento = janela.child_window(handle=controle.handle)

    trazer_janela_para_frente(hwnd_janela)
    try:
        elemento.set_focus()
    except Exception:  # noqa: BLE001
        pass

    for acao in (
        lambda: elemento.click(),
        lambda: elemento.type_keys("{SPACE}"),
        lambda: elemento.click_input(),
    ):
        try:
            acao()
            return
        except Exception:  # noqa: BLE001
            continue

    raise RuntimeError(f"Nao foi possivel acionar o controle '{controle.texto}'.")


def clicar_botao_por_texto(hwnd_janela: int, trechos: Sequence[str]) -> bool:
    controle = localizar_controle_por_texto(hwnd_janela, trechos, classes_permitidas=("button",))
    if controle is None:
        return False

    acionar_controle(hwnd_janela, controle)
    return True


def garantir_checkbox_marcado(hwnd_janela: int, trechos: Sequence[str]) -> bool:
    controle = localizar_controle_por_texto(hwnd_janela, trechos, classes_permitidas=("check", "button"))
    if controle is None:
        return False

    try:
        estado = win32gui.SendMessage(controle.handle, win32con.BM_GETCHECK, 0, 0)
    except Exception:  # noqa: BLE001
        estado = None

    if estado in (win32con.BST_CHECKED, 1):
        return True

    acionar_controle(hwnd_janela, controle)
    return True


def pressionar_enter_na_janela(
    hwnd_janela: int,
    vezes: int,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> None:
    trazer_janela_para_frente(hwnd_janela)
    janela = obter_wrapper_janela(hwnd_janela)

    for indice in range(1, vezes + 1):
        verificar_cancelamento(cancelar_evento, "envio de Enter para o instalador")
        try:
            janela.set_focus()
        except Exception:  # noqa: BLE001
            pass

        enviar_teclas_pywinauto("{ENTER}")
        logar(logger, f"Enter {indice}/{vezes} enviado para o instalador.")
        dormir_interrompivel(
            TEMPO_ESPERA_ENTRE_ENTERS,
            cancelar_evento,
            "espera entre os Enters do instalador",
        )


def localizar_executavel_interface(pasta_interface: Path) -> Path:
    preferidos = (
        pasta_interface / "InterfaceSi.exe",
        pasta_interface / "Interface.exe",
    )

    for candidato in preferidos:
        if candidato.exists():
            return candidato

    executaveis = sorted(
        arquivo
        for arquivo in pasta_interface.glob("*.exe")
        if arquivo.is_file() and "interface" in arquivo.name.lower() and not arquivo.name.lower().startswith("unins")
    )
    if executaveis:
        return executaveis[0]

    raise RuntimeError(f"Nenhum executavel principal da Interface foi encontrado em {pasta_interface}.")


def coletar_processos_por_caminho(caminho_executavel: Path) -> list[psutil.Process]:
    encontrados: list[psutil.Process] = []
    caminho_desejado = str(caminho_executavel.resolve()).lower()

    for processo in psutil.process_iter(["pid", "exe", "name"]):
        try:
            exe = processo.info.get("exe")
            if exe and str(Path(exe).resolve()).lower() == caminho_desejado:
                encontrados.append(processo)
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            continue

    return encontrados


def coletar_pids_interface_novos(caminho_executavel: Path, pids_anteriores: set[int]) -> set[int]:
    return {
        processo.pid
        for processo in coletar_processos_por_caminho(caminho_executavel)
        if processo.pid not in pids_anteriores
    }


def pid_corresponde_ao_executavel(pid: int, caminho_executavel: Path) -> bool:
    try:
        processo = psutil.Process(pid)
        exe = processo.exe()
    except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
        return False
    return str(Path(exe).resolve()).lower() == str(caminho_executavel.resolve()).lower()


def localizar_janela_preferencial_interface(pids: Iterable[int]) -> JanelaDetectada | None:
    janelas = listar_janelas_visiveis(pids)
    if not janelas:
        return None

    def prioridade(janela: JanelaDetectada) -> tuple[int, int, str]:
        titulo = normalizar_texto_ui(janela.titulo)
        classe = normalizar_texto_ui(janela.classe)
        if "acesso ao sistema" in titulo or "login" in titulo or "senha" in classe:
            grupo = 0
        elif "interface 1.0" in titulo or classe == "tprincipal":
            grupo = 1
        elif classe == "tapplication":
            grupo = 3
        else:
            grupo = 2
        return (grupo, 0 if janela.titulo else 1, titulo)

    return sorted(janelas, key=prioridade)[0]


def localizar_tela_acesso_interface(
    caminho_executavel: Path,
    pids_anteriores: set[int],
) -> JanelaDetectada | None:
    pids_novos = coletar_pids_interface_novos(caminho_executavel, pids_anteriores)
    candidatas = [
        janela
        for janela in listar_todas_janelas_visiveis()
        if "acesso ao sistema" in normalizar_texto_ui(janela.titulo)
    ]
    if not candidatas:
        return None

    def prioridade(janela: JanelaDetectada) -> tuple[int, int, int, str]:
        return (
            0 if janela.pid in pids_novos else 1,
            0 if pid_corresponde_ao_executavel(janela.pid, caminho_executavel) else 1,
            0 if janela.titulo else 1,
            normalizar_texto_ui(janela.titulo),
        )

    return sorted(candidatas, key=prioridade)[0]


def tentar_fechar_janela_interface(
    janela: JanelaDetectada,
    logger: Callable[[str], None] | None = None,
) -> bool:
    titulo = janela.titulo or janela.classe
    logar(logger, f"Tentando fechar a Interface na janela '{titulo}'.")

    try:
        if clicar_botao_por_texto(janela.handle, ("cancelar", "fechar", "sair")):
            logar(logger, "Botao de fechamento encontrado e acionado.")
            return True
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Falha ao clicar no botao de fechamento: {exc}")

    trazer_janela_para_frente(janela.handle)

    try:
        wrapper = obter_wrapper_janela(janela.handle)
    except Exception:  # noqa: BLE001
        wrapper = None

    for descricao, teclas in (
        ("atalho ESC", "{ESC}"),
        ("atalho Alt+F4", "%{F4}"),
    ):
        try:
            if wrapper is not None:
                wrapper.set_focus()
                wrapper.type_keys(teclas, set_foreground=True)
            else:
                enviar_teclas_pywinauto(teclas)
            logar(logger, f"{descricao} enviado para a Interface.")
            time.sleep(TEMPO_ESPERA_REACAO_FECHAMENTO)
            if not janela_continua_visivel(janela.handle):
                return True
            logar(logger, f"{descricao} nao fechou a Interface; tentando o proximo fallback.")
        except Exception as exc:  # noqa: BLE001
            logar(logger, f"Nao foi possivel usar {descricao}: {exc}")

    titulo_normalizado = normalizar_texto_ui(janela.titulo)
    if "acesso ao sistema" in titulo_normalizado:
        try:
            clicar_cancelar_interface(janela, logger=logger)
            time.sleep(TEMPO_ESPERA_REACAO_FECHAMENTO)
            if not janela_continua_visivel(janela.handle):
                logar(logger, "Tela de acesso fechada pelo fallback visual do botao Cancelar.")
                return True
        except Exception as exc:  # noqa: BLE001
            logar(logger, f"Falha no fallback visual do botao Cancelar: {exc}")

    try:
        win32gui.PostMessage(janela.handle, win32con.WM_CLOSE, 0, 0)
        logar(logger, "WM_CLOSE enviado para a janela da Interface.")
        return True
    except Exception as exc:  # noqa: BLE001
        logar(logger, f"Nao foi possivel enviar WM_CLOSE: {exc}")

    return False


def encerrar_processos_por_pid(
    pids: Iterable[int],
    descricao: str,
    logger: Callable[[str], None] | None = None,
) -> None:
    pids_ativos = sorted({int(pid) for pid in pids if pid and psutil.pid_exists(int(pid))})
    if not pids_ativos:
        return

    logar(logger, f"Encerrando {descricao} nos PID(s): {pids_ativos}")
    for pid in pids_ativos:
        try:
            psutil.Process(pid).terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    prazo = time.monotonic() + 10
    while time.monotonic() < prazo:
        restantes = [pid for pid in pids_ativos if psutil.pid_exists(pid)]
        if not restantes:
            return
        time.sleep(0.5)

    for pid in pids_ativos:
        try:
            psutil.Process(pid).kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def encerrar_processos_interface_restantes(
    caminho_executavel: Path,
    pids_anteriores: set[int],
    logger: Callable[[str], None] | None = None,
) -> None:
    pids_ativos = coletar_pids_interface_novos(caminho_executavel, pids_anteriores)
    if not pids_ativos:
        return

    encerrar_processos_por_pid(pids_ativos, "a Interface", logger=logger)


def aguardar_encerramento_dos_pids(
    pids: Iterable[int],
    timeout: float,
    cancelar_evento: threading.Event | None = None,
    contexto: str = "encerramento da Interface",
) -> bool:
    pids_monitorados = {int(pid) for pid in pids if pid}
    prazo = time.monotonic() + max(0.0, timeout)

    while time.monotonic() < prazo:
        verificar_cancelamento(cancelar_evento, contexto)
        if pids_monitorados:
            pids_monitorados.update(coletar_pids_relacionados(pids_monitorados))

        ativos = [pid for pid in pids_monitorados if psutil.pid_exists(pid)]
        if not ativos:
            return True

        dormir_interrompivel(0.4, cancelar_evento, contexto)

    return not any(psutil.pid_exists(pid) for pid in pids_monitorados)


def garantir_interface_encerrada_por_pids(
    pids_interface: Iterable[int],
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> bool:
    pids_monitorados = {int(pid) for pid in pids_interface if pid}
    if not pids_monitorados:
        return True

    for tentativa in range(1, 4):
        verificar_cancelamento(cancelar_evento, "encerramento final da Interface")
        pids_monitorados.update(coletar_pids_relacionados(pids_monitorados))
        pids_ativos = [pid for pid in sorted(pids_monitorados) if psutil.pid_exists(pid)]
        if not pids_ativos:
            return True

        janelas = listar_janelas_visiveis(pids_ativos)
        if janelas:
            logar(
                logger,
                "A Interface ainda possui janela(s) aberta(s); "
                f"tentando fechar antes de seguir para o proximo setup (tentativa {tentativa}).",
            )
            for janela in janelas:
                try:
                    tentar_fechar_janela_interface(janela, logger=logger)
                except Exception as exc:  # noqa: BLE001
                    titulo = janela.titulo or janela.classe or str(janela.handle)
                    logar(logger, f"Falha ao fechar a janela remanescente '{titulo}': {exc}")
        else:
            logar(
                logger,
                "A Interface ainda possui processo(s) sem janela visivel; "
                f"aguardando o encerramento completo (tentativa {tentativa}).",
            )

        if aguardar_encerramento_dos_pids(
            pids_monitorados,
            TEMPO_ESPERA_FECHAMENTO_INTERFACE + 2,
            cancelar_evento=cancelar_evento,
            contexto="espera pelo encerramento completo da Interface",
        ):
            return True

    pids_finais = [pid for pid in sorted(pids_monitorados) if psutil.pid_exists(pid)]
    if not pids_finais:
        return True

    logar(
        logger,
        "A Interface permaneceu aberta apos as tentativas de fechamento visual; "
        f"encerrando PID(s): {pids_finais}",
    )
    encerrar_processos_por_pid(pids_finais, "a Interface", logger=logger)
    return aguardar_encerramento_dos_pids(
        pids_monitorados,
        5,
        cancelar_evento=cancelar_evento,
        contexto="confirmacao final do encerramento da Interface",
    )


def garantir_processos_interface_encerrados(
    executavel_interface: Path,
    pids_anteriores: set[int],
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> bool:
    prazo = time.monotonic() + max(8.0, TEMPO_ESPERA_FECHAMENTO_INTERFACE + 5.0)
    tentativa = 0

    while time.monotonic() < prazo:
        verificar_cancelamento(cancelar_evento, "encerramento completo da Interface")
        restantes = coletar_pids_interface_novos(executavel_interface, pids_anteriores)
        if not restantes:
            logar(logger, "Interface encerrada por completo antes de seguir para a proxima atualizacao.")
            return True

        tentativa += 1
        janelas = listar_janelas_visiveis(restantes)
        if janelas:
            logar(
                logger,
                "A Interface ainda esta aberta apos fechar a tela de acesso; "
                f"tentando encerrar as janelas remanescentes (tentativa {tentativa}, PID(s) {sorted(restantes)}).",
            )
            for janela in janelas:
                try:
                    tentar_fechar_janela_interface(janela, logger=logger)
                except Exception as exc:  # noqa: BLE001
                    titulo = janela.titulo or janela.classe or str(janela.handle)
                    logar(logger, f"Falha ao fechar a janela remanescente '{titulo}': {exc}")
        elif tentativa == 1:
            logar(
                logger,
                "A Interface ainda possui processo(s) em segundo plano apos fechar a tela de acesso; "
                f"aguardando o encerramento completo dos PID(s) {sorted(restantes)}.",
            )

        dormir_interrompivel(1.0, cancelar_evento, "espera pelo encerramento completo da Interface")

    restantes = coletar_pids_interface_novos(executavel_interface, pids_anteriores)
    if not restantes:
        logar(logger, "Interface encerrada por completo antes de seguir para a proxima atualizacao.")
        return True

    logar(
        logger,
        "A Interface permaneceu aberta apos o fechamento da tela de acesso; "
        f"aplicando encerramento forcado nos PID(s) {sorted(restantes)}.",
    )
    encerrar_processos_por_pid(restantes, "a Interface", logger=logger)

    restantes_finais = coletar_pids_interface_novos(executavel_interface, pids_anteriores)
    if not restantes_finais:
        logar(logger, "Interface encerrada por completo antes de seguir para a proxima atualizacao.")
        return True

    logar(
        logger,
        "Ainda restaram processo(s) da Interface apos o encerramento forcado: "
        f"{sorted(restantes_finais)}.",
    )
    return False


def clicar_cancelar_interface(
    janela_login: JanelaDetectada,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> None:
    verificar_cancelamento(cancelar_evento, "fechamento da tela de acesso com ESC")
    logar(logger, "ESC sera enviado diretamente para fechar a tela de acesso da Interface.")
    if enviar_escape_para_janela(janela_login.handle, logger=logger):
        dormir_interrompivel(
            TEMPO_ESPERA_REACAO_FECHAMENTO,
            cancelar_evento,
            "espera apos ESC na tela de acesso",
        )


def calcular_tamanho_inicial_janela(
    janela: tk.Misc,
    largura_desejada: int,
    altura_desejada: int,
    margem_horizontal: int = 120,
    margem_vertical: int = 120,
) -> tuple[int, int]:
    largura_tela = max(1, janela.winfo_screenwidth())
    altura_tela = max(1, janela.winfo_screenheight())
    largura_maxima = max(320, largura_tela - margem_horizontal)
    altura_maxima = max(240, altura_tela - margem_vertical)
    return (
        max(320, min(largura_desejada, largura_maxima)),
        max(240, min(altura_desejada, altura_maxima)),
    )


def aplicar_geometria_inicial(
    janela: tk.Misc,
    largura_desejada: int,
    altura_desejada: int,
    min_largura: int,
    min_altura: int,
) -> None:
    largura, altura = calcular_tamanho_inicial_janela(janela, largura_desejada, altura_desejada)
    x = max((janela.winfo_screenwidth() - largura) // 2, 0)
    y = max((janela.winfo_screenheight() - altura) // 2, 0)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    janela.minsize(min(min_largura, largura), min(min_altura, altura))


def aguardar_cancelar_e_fechar_interface(
    executavel_interface: Path,
    pids_anteriores: set[int],
    janela_login_inicial: JanelaDetectada | None = None,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> bool:
    prazo = time.monotonic() + TEMPO_LIMITE_LOCALIZAR_CANCELAR
    tentativa = 0

    while time.monotonic() < prazo:
        verificar_cancelamento(cancelar_evento, "espera pela tela de acesso da Interface")
        janela_login = None
        if janela_login_inicial is not None and janela_continua_visivel(janela_login_inicial.handle):
            janela_login = janela_login_inicial
        else:
            restantes = coletar_pids_interface_novos(executavel_interface, pids_anteriores)
            if not restantes:
                if janela_login_inicial is not None and not janela_continua_visivel(janela_login_inicial.handle):
                    logar(logger, "Tela de acesso fechada com sucesso apos localizar o Cancelar.")
                    return True
                time.sleep(0.5)
                continue

            janela_login = localizar_janela_preferencial_interface(restantes)
            if janela_login is None:
                time.sleep(0.5)
                continue

        titulo_normalizado = normalizar_texto_ui(janela_login.titulo)
        if "acesso ao sistema" not in titulo_normalizado:
            time.sleep(0.5)
            continue

        tentativa += 1
        logar(logger, f"Tela de acesso detectada. Procurando o Cancelar (tentativa {tentativa}).")

        try:
            clicar_cancelar_interface(janela_login, logger=logger, cancelar_evento=cancelar_evento)
        except Exception as exc:  # noqa: BLE001
            logar(logger, f"Falha ao tentar localizar/clicar em Cancelar: {exc}")

        limite_reacao = time.monotonic() + TEMPO_ESPERA_FECHAMENTO_INTERFACE
        while time.monotonic() < limite_reacao:
            verificar_cancelamento(cancelar_evento, "espera pelo fechamento da Interface")
            if not janela_continua_visivel(janela_login.handle):
                logar(logger, "Tela de acesso fechada apos clicar em Cancelar.")
                return True
            restantes = coletar_pids_interface_novos(executavel_interface, pids_anteriores)
            if not restantes:
                logar(logger, "Interface fechada com sucesso apos clicar em Cancelar.")
                return True
            dormir_interrompivel(0.4, cancelar_evento, "espera pelo fechamento da Interface")

        logar(logger, "A Interface ainda esta aberta; vou procurar o Cancelar novamente.")
        dormir_interrompivel(0.6, cancelar_evento, "nova tentativa de localizar o Cancelar")

    return not coletar_pids_interface_novos(executavel_interface, pids_anteriores)


def aguardar_interface_abrir_e_fechar(
    pasta_interface: Path,
    pids_anteriores: set[int] | None = None,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> None:
    executavel_interface = localizar_executavel_interface(pasta_interface)
    pids_antes = set(pids_anteriores or ())
    if not pids_antes:
        pids_antes = {processo.pid for processo in coletar_processos_por_caminho(executavel_interface)}

    janela_login = aguardar_condicao(
        "tela de acesso da Interface",
        lambda: localizar_tela_acesso_interface(executavel_interface, pids_antes),
        TEMPO_LIMITE_ABERTURA_INTERFACE,
        intervalo=1.0,
        cancelar_evento=cancelar_evento,
    )
    assert isinstance(janela_login, JanelaDetectada)
    logar(logger, f"Tela de acesso da Interface detectada (PID {janela_login.pid}). Iniciando busca pelo Cancelar.")
    fechou_tela_acesso = aguardar_cancelar_e_fechar_interface(
        executavel_interface,
        pids_antes,
        janela_login_inicial=janela_login,
        logger=logger,
        cancelar_evento=cancelar_evento,
    )
    if not fechou_tela_acesso:
        logar(logger, "A Interface nao fechou pela tela de acesso dentro do prazo. Aplicando encerramento de seguranca.")

    if garantir_processos_interface_encerrados(
        executavel_interface,
        pids_antes,
        logger=logger,
        cancelar_evento=cancelar_evento,
    ):
        return

    restantes = coletar_pids_interface_novos(executavel_interface, pids_antes)
    if restantes:
        raise RuntimeError(
            "A Interface continuou aberta apos clicar em Cancelar: "
            f"PID(s) remanescentes {sorted(restantes)}"
        )
    logar(logger, "Interface encerrada por fallback apos o clique em Cancelar.")


def automatizar_setup(
    setup: SetupAtualizacao,
    pasta_interface: Path,
    argumentos: Sequence[str],
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> None:
    verificar_cancelamento(cancelar_evento, f"inicio do setup {setup.caminho.name}")
    with bloquear_entrada_usuario(logger, f"a automacao do setup {setup.caminho.name}"):
        executavel_interface = localizar_executavel_interface(pasta_interface)
        pids_interface_antes = {processo.pid for processo in coletar_processos_por_caminho(executavel_interface)}
        processo_raiz = subprocess.Popen([str(setup.caminho), *argumentos], cwd=str(setup.caminho.parent))
        pids_monitorados: set[int] = {processo_raiz.pid}
        logar(logger, f"Setup iniciado com PID {processo_raiz.pid}.")

        def atualizar_pids() -> list[int]:
            pids_monitorados.update(coletar_pids_relacionados(pids_monitorados))
            return sorted(pids_monitorados)

        try:
            janela_idioma = aguardar_janela(
                atualizar_pids,
                "janela de idioma do instalador",
                TEMPO_LIMITE_JANELA_INSTALADOR,
                lambda janela: "selecionar idioma" in normalizar_texto_ui(janela.titulo),
                cancelar_evento=cancelar_evento,
            )
            logar(logger, "Janela de idioma localizada. Confirmando idioma padrao.")
            if not clicar_botao_por_texto(janela_idioma.handle, ("ok",)):
                raise RuntimeError("Nao foi possivel confirmar a janela de idioma do instalador.")

            janela_principal = aguardar_janela(
                atualizar_pids,
                "janela principal do instalador",
                TEMPO_LIMITE_JANELA_INSTALADOR,
                lambda janela: janela.classe != "TApplication" and "selecionar idioma" not in normalizar_texto_ui(janela.titulo),
                cancelar_evento=cancelar_evento,
            )
            pressionar_enter_na_janela(
                janela_principal.handle,
                ENTERS_POR_SETUP,
                logger=logger,
                cancelar_evento=cancelar_evento,
            )

            def localizar_janela_conclusao() -> JanelaDetectada | None:
                for janela in listar_janelas_visiveis(atualizar_pids()):
                    if janela.classe == "TApplication":
                        continue

                    controle_checkbox = localizar_controle_por_texto(
                        janela.handle,
                        ("executar", "interface", "run"),
                        classes_permitidas=("check", "button"),
                    )
                    controle_finalizar = localizar_controle_por_texto(
                        janela.handle,
                        ("concluir", "finalizar", "finish"),
                        classes_permitidas=("button",),
                    )
                    if controle_checkbox or controle_finalizar:
                        return janela
                return None

            janela_conclusao = aguardar_condicao(
                "pagina final do instalador",
                localizar_janela_conclusao,
                TEMPO_LIMITE_CONCLUSAO_INSTALADOR,
                intervalo=1.0,
                cancelar_evento=cancelar_evento,
            )
            assert isinstance(janela_conclusao, JanelaDetectada)
            logar(logger, "Pagina final do instalador localizada.")

            if garantir_checkbox_marcado(janela_conclusao.handle, ("executar interface", "interface")):
                logar(logger, "Opcao para executar a Interface ao final confirmada.")
            else:
                logar(logger, "Nao foi possivel confirmar visualmente o checkbox da Interface; seguindo com a finalizacao.")

            if clicar_botao_por_texto(janela_conclusao.handle, ("concluir", "finalizar", "finish")):
                logar(logger, "Botao de conclusao acionado.")
            else:
                trazer_janela_para_frente(janela_conclusao.handle)
                enviar_teclas_pywinauto("{ENTER}")
                logar(logger, "Botao de conclusao nao foi localizado; Enter enviado como alternativa.")

            aguardar_condicao(
                "encerramento do instalador",
                lambda: all(not psutil.pid_exists(pid) for pid in atualizar_pids()),
                120,
                intervalo=1.0,
                cancelar_evento=cancelar_evento,
            )
            logar(logger, "Instalador encerrado.")

            aguardar_interface_abrir_e_fechar(
                pasta_interface,
                pids_anteriores=pids_interface_antes,
                logger=logger,
                cancelar_evento=cancelar_evento,
            )
        except OperacaoCancelada:
            logar(logger, f"Parada solicitada durante o setup {setup.caminho.name}.")
            encerrar_processos_por_pid(atualizar_pids(), "o instalador em execucao", logger=logger)
            encerrar_processos_interface_restantes(executavel_interface, pids_interface_antes, logger=logger)
            raise


def configurar_cliente_firebird(pasta_interface: Path) -> None:
    fbclient = pasta_interface / "fbclient.dll"
    if fbclient.exists():
        try:
            fdb.load_api(str(fbclient))
        except OSError:
            pass


def coletar_arquivos_ini(pasta_interface: Path) -> list[Path]:
    arquivos: list[Path] = []
    for pasta in (pasta_interface, pasta_interface / "Bd"):
        if not pasta.exists():
            continue
        arquivos.extend(sorted(pasta.glob("*.ini")))
    return arquivos


def caminho_banco_deve_ser_ignorado(caminho: Path) -> bool:
    texto_caminho = normalizar_nome_arquivo(str(caminho).replace("\\", " "))
    nome_arquivo = normalizar_nome_arquivo(caminho.name)
    return any(
        indicador in texto_caminho or indicador in nome_arquivo
        for indicador in INDICADORES_CAMINHO_BANCO_IGNORADO
    )


def ler_candidatos_dos_ini(arquivos_ini: Iterable[Path]) -> list[BancoCandidato]:
    candidatos: list[BancoCandidato] = []

    for arquivo_ini in arquivos_ini:
        parser = configparser.ConfigParser()
        parser.optionxform = str

        try:
            parser.read(arquivo_ini, encoding="latin-1")
        except configparser.Error:
            continue

        for secao in parser.sections():
            caminho_banco = (
                parser.get(secao, "Database", fallback="")
                or parser.get(secao, "DataBase", fallback="")
                or parser.get(secao, "CaminhoBancoDados", fallback="")
            ).strip()

            if not caminho_banco:
                continue

            caminho = Path(caminho_banco)
            if caminho.suffix.lower() not in EXTENSOES_FIREBIRD:
                continue
            if caminho_banco_deve_ser_ignorado(caminho):
                continue

            candidatos.append(
                BancoCandidato(
                    caminho=caminho,
                    usuario=parser.get(secao, "User_Name", fallback="SYSDBA").strip() or "SYSDBA",
                    senha=parser.get(secao, "Password", fallback="masterkey").strip(),
                    origem=f"{arquivo_ini.name} [{secao}]",
                    referenciado_em_config=True,
                )
            )

    return candidatos


def ler_candidatos_por_varredura(pasta_interface: Path) -> list[BancoCandidato]:
    candidatos: list[BancoCandidato] = []
    for arquivo in sorted(pasta_interface.rglob("*")):
        if not arquivo.is_file():
            continue
        if arquivo.suffix.lower() not in EXTENSOES_FIREBIRD:
            continue
        if caminho_banco_deve_ser_ignorado(arquivo):
            continue
        candidatos.append(BancoCandidato(caminho=arquivo, origem="varredura da pasta"))
    return candidatos


def ordenar_candidatos(candidatos: Iterable[BancoCandidato], pasta_interface: Path) -> list[BancoCandidato]:
    unicos: dict[str, BancoCandidato] = {}
    referencias_config: dict[str, int] = {}

    for candidato in candidatos:
        if caminho_banco_deve_ser_ignorado(candidato.caminho):
            continue
        chave = str(candidato.caminho).lower()
        if candidato.referenciado_em_config:
            referencias_config[chave] = referencias_config.get(chave, 0) + 1

        existente = unicos.get(chave)
        if existente is None:
            unicos[chave] = candidato
            continue
        if candidato.referenciado_em_config and not existente.referenciado_em_config:
            unicos[chave] = candidato

    pasta_bd = str((pasta_interface / "Bd")).lower()

    def prioridade(candidato: BancoCandidato) -> tuple[int, int, int, int, int, str]:
        caminho_texto = str(candidato.caminho).lower()
        existe = candidato.caminho.exists()
        referencias = referencias_config.get(caminho_texto, 0)
        esta_na_pasta_bd = 0 if pasta_bd in caminho_texto else 1
        eh_base_isolada = 1 if candidato.caminho.name.lower() == NOME_BASE_ISOLADA.lower() else 0
        eh_base_interface = 0 if "bd_interface" in candidato.caminho.name.lower() else 1
        return (
            0 if existe else 1,
            0 if referencias > 0 else 1,
            -referencias,
            esta_na_pasta_bd,
            eh_base_isolada,
            eh_base_interface,
            caminho_texto,
        )

    return sorted(unicos.values(), key=prioridade)


def encontrar_tabela_compilacao(cursor: fdb.Cursor) -> str | None:
    cursor.execute(
        """
        SELECT TRIM(rf.RDB$RELATION_NAME)
        FROM RDB$RELATION_FIELDS rf
        JOIN RDB$RELATIONS r
          ON r.RDB$RELATION_NAME = rf.RDB$RELATION_NAME
        WHERE COALESCE(r.RDB$SYSTEM_FLAG, 0) = 0
          AND UPPER(TRIM(rf.RDB$FIELD_NAME)) = ?
        """,
        (CAMPO_COMPILACAO,),
    )

    tabelas = [linha[0].strip() for linha in cursor.fetchall()]
    if not tabelas:
        return None

    preferencia = {"T000": 0, "T000_CONFIGURACOES": 1}
    tabelas.sort(key=lambda nome: (preferencia.get(nome.upper(), 99), nome))
    return tabelas[0]


def abrir_conexao_banco(candidato: BancoCandidato):
    return fdb.connect(
        dsn=str(candidato.caminho),
        user=candidato.usuario or "SYSDBA",
        password=candidato.senha or "masterkey",
        charset="WIN1252",
    )


def normalizar_versao_banco(valor: object) -> VersaoBanco:
    if isinstance(valor, datetime):
        return VersaoBanco(valor_bruto=valor, compilado_em=valor, tem_hora=True)
    if isinstance(valor, date):
        return VersaoBanco(
            valor_bruto=valor,
            compilado_em=datetime(valor.year, valor.month, valor.day),
            tem_hora=False,
        )
    raise RuntimeError(f"Valor inesperado em {CAMPO_COMPILACAO}: {valor!r}")


def formatar_versao(versao: VersaoBanco) -> str:
    mascara = "%d/%m/%Y %H:%M" if versao.tem_hora else "%d/%m/%Y"
    return versao.compilado_em.strftime(mascara)


def ler_ts_compilacao(candidato: BancoCandidato) -> tuple[str, VersaoBanco]:
    conexao = abrir_conexao_banco(candidato)

    try:
        cursor = conexao.cursor()
        tabela = encontrar_tabela_compilacao(cursor)
        if not tabela:
            raise RuntimeError(f"O banco nao possui o campo {CAMPO_COMPILACAO}.")

        cursor.execute(f"SELECT FIRST 1 {CAMPO_COMPILACAO} FROM {tabela}")
        linha = cursor.fetchone()
        valor = linha[0] if linha else None
        if valor is None:
            raise RuntimeError(f"O campo {CAMPO_COMPILACAO} esta vazio.")
        return tabela, normalizar_versao_banco(valor)
    finally:
        conexao.close()


def carregar_script_sql(caminho_script: Path) -> str:
    with caminho_script.open("r", encoding="utf-8") as arquivo:
        return arquivo.read()


def dividir_comandos_sql(script_sql: str) -> list[str]:
    comandos: list[str] = []
    atual: list[str] = []
    em_texto = False
    delimitador = ""

    for caractere in script_sql:
        if caractere in {"'", '"'}:
            if em_texto and caractere == delimitador:
                em_texto = False
                delimitador = ""
            elif not em_texto:
                em_texto = True
                delimitador = caractere
            atual.append(caractere)
            continue

        if caractere == ";" and not em_texto:
            comando = "".join(atual).strip()
            if comando:
                comandos.append(comando)
            atual = []
            continue

        atual.append(caractere)

    comando_final = "".join(atual).strip()
    if comando_final:
        comandos.append(comando_final)
    return comandos


def aplicar_correcao_grid_localizacao_produtos(
    analise: ResultadoAnalise,
    logger: Callable[[str], None] | None = None,
) -> None:
    caminho_script = ARQUIVO_SQL_CORRECAO_GRID
    if not caminho_script.exists():
        raise RuntimeError(f"Arquivo SQL de correcao da grid nao encontrado: {caminho_script}")

    script_sql = carregar_script_sql(caminho_script)
    comandos = dividir_comandos_sql(script_sql)
    if not comandos:
        raise RuntimeError(f"O arquivo SQL de correcao da grid esta vazio: {caminho_script}")

    logar(
        logger,
        f"Aplicando correcao da grid de localizacao de produtos em {analise.banco_principal.caminho.name}.",
    )
    conexao = abrir_conexao_banco(analise.banco_principal)

    try:
        cursor = conexao.cursor()
        executados = 0
        for comando in comandos:
            normalizado = re.sub(r"\s+", " ", comando).strip().casefold()
            if not normalizado:
                continue

            if normalizado.startswith("commit"):
                conexao.commit()
                continue

            if normalizado.startswith("rollback"):
                conexao.rollback()
                continue

            cursor.execute(comando)
            executados += 1

        conexao.commit()
        logar(logger, f"Correcao da grid concluida com {executados} comando(s) SQL executado(s).")
    except Exception:
        conexao.rollback()
        raise
    finally:
        conexao.close()


def localizar_banco_principal(pasta_interface: Path) -> tuple[BancoCandidato, str, VersaoBanco]:
    configurar_cliente_firebird(pasta_interface)

    arquivos_ini = coletar_arquivos_ini(pasta_interface)
    candidatos = ordenar_candidatos(
        [*ler_candidatos_dos_ini(arquivos_ini), *ler_candidatos_por_varredura(pasta_interface)],
        pasta_interface,
    )

    if not candidatos:
        raise RuntimeError(f"Nenhum banco Firebird foi encontrado em {pasta_interface}.")

    erros: list[str] = []

    for candidato in candidatos:
        if not candidato.caminho.exists():
            erros.append(f"{candidato.caminho} -> caminho nao encontrado ({candidato.origem})")
            continue

        try:
            tabela, valor = ler_ts_compilacao(candidato)
        except Exception as exc:  # noqa: BLE001
            erros.append(f"{candidato.caminho} -> {exc}")
            continue

        return candidato, tabela, valor

    detalhes = "\n".join(f"- {erro}" for erro in erros)
    raise RuntimeError(f"Nao foi possivel localizar um banco valido em {pasta_interface}.\n{detalhes}")


def extrair_data_setup(caminho: Path) -> datetime | None:
    correspondencia = re.search(r"(\d{12})(?!\d)", caminho.stem)
    if not correspondencia:
        return None

    try:
        return datetime.strptime(correspondencia.group(1), "%d%m%Y%H%M")
    except ValueError:
        return None


def resolver_pasta_setups(pasta_setups: Path) -> Path:
    if pasta_setups.exists():
        return pasta_setups

    bases_busca: list[Path] = []
    vistos: set[str] = set()

    def adicionar_bases(origem: Path) -> None:
        for base in listar_bases_busca_setups(origem):
            chave = str(base.resolve()).lower()
            if chave in vistos:
                continue
            vistos.add(chave)
            bases_busca.append(base)

    if str(pasta_setups.parent) not in {"", "."}:
        adicionar_bases(pasta_setups.parent)
    adicionar_bases(PASTA_APLICACAO)
    adicionar_bases(Path.cwd())

    for base_dir in bases_busca:
        encontrada = encontrar_pasta_setups_semelhante(base_dir, pasta_setups.name)
        if encontrada is not None:
            return encontrada

    raise RuntimeError(f"Pasta de setups nao encontrada: {pasta_setups}")


def coletar_setups(pasta_setups: Path) -> list[SetupAtualizacao]:
    pasta_setups = resolver_pasta_setups(pasta_setups)
    setups: list[SetupAtualizacao] = []
    for arquivo in sorted(pasta_setups.glob("*.exe")):
        data_setup = extrair_data_setup(arquivo)
        if data_setup is None:
            continue
        setups.append(SetupAtualizacao(caminho=arquivo, compilado_em=data_setup))
    return sorted(setups, key=lambda item: item.compilado_em)


def setup_eh_mais_novo(versao_atual: VersaoBanco, setup: SetupAtualizacao) -> bool:
    if versao_atual.tem_hora:
        return setup.compilado_em > versao_atual.compilado_em
    return setup.compilado_em.date() > versao_atual.compilado_em.date()


def selecionar_setups_mes_a_mes(
    setups: Sequence[SetupAtualizacao],
    versao_atual: VersaoBanco,
) -> list[SetupAtualizacao]:
    por_mes: dict[tuple[int, int], SetupAtualizacao] = {}

    for setup in setups:
        if not setup_eh_mais_novo(versao_atual, setup):
            continue

        chave = (setup.compilado_em.year, setup.compilado_em.month)
        escolhido = por_mes.get(chave)
        if escolhido is None or setup.compilado_em > escolhido.compilado_em:
            por_mes[chave] = setup

    return [por_mes[chave] for chave in sorted(por_mes)]


def analisar_atualizacao(pasta_interface: Path, pasta_setups: Path) -> ResultadoAnalise:
    banco_principal, tabela_compilacao, versao_atual = localizar_banco_principal(pasta_interface)
    pasta_setups = resolver_pasta_setups(pasta_setups)
    setups_disponiveis = coletar_setups(pasta_setups)
    fila_atualizacao = selecionar_setups_mes_a_mes(setups_disponiveis, versao_atual)

    return ResultadoAnalise(
        pasta_interface=pasta_interface,
        pasta_setups=pasta_setups,
        banco_principal=banco_principal,
        tabela_compilacao=tabela_compilacao,
        versao_atual=versao_atual,
        fila_atualizacao=fila_atualizacao,
        setups_disponiveis=setups_disponiveis,
        base_isolada=banco_principal.caminho.name.lower() == NOME_BASE_ISOLADA.lower(),
    )


def ler_texto_preservando_quebras(caminho: Path) -> str:
    with caminho.open("r", encoding="latin-1", errors="ignore", newline="") as arquivo:
        return arquivo.read()


def escrever_texto_preservando_quebras(caminho: Path, conteudo: str) -> None:
    with caminho.open("w", encoding="latin-1", newline="") as arquivo:
        arquivo.write(conteudo)


def planejar_alteracoes_configuracao(
    pasta_interface: Path,
    nome_atual: str,
    nome_novo: str,
) -> list[AlteracaoArquivo]:
    padrao = re.compile(re.escape(nome_atual), re.IGNORECASE)
    alteracoes: list[AlteracaoArquivo] = []

    for arquivo in sorted(pasta_interface.rglob("*.ini")):
        conteudo = ler_texto_preservando_quebras(arquivo)
        novo_conteudo, ocorrencias = padrao.subn(nome_novo, conteudo)
        if ocorrencias:
            alteracoes.append(
                AlteracaoArquivo(
                    caminho=arquivo,
                    conteudo_original=conteudo,
                    conteudo_novo=novo_conteudo,
                    ocorrencias=ocorrencias,
                )
            )

    return alteracoes


def criar_backup_unico(caminho: Path) -> Path:
    carimbo = datetime.now().strftime("%Y%m%d%H%M%S")
    backup = caminho.with_name(f"{caminho.name}.{carimbo}.bak")
    indice = 1
    while backup.exists():
        backup = caminho.with_name(f"{caminho.name}.{carimbo}.{indice}.bak")
        indice += 1
    shutil.copy2(caminho, backup)
    return backup


def aplicar_alteracoes_configuracao(
    alteracoes: Sequence[AlteracaoArquivo],
    logger: Callable[[str], None] | None = None,
) -> list[tuple[Path, Path]]:
    backups: list[tuple[Path, Path]] = []
    for alteracao in alteracoes:
        backup = criar_backup_unico(alteracao.caminho)
        backups.append((alteracao.caminho, backup))
        escrever_texto_preservando_quebras(alteracao.caminho, alteracao.conteudo_novo)
        if logger is not None:
            logger(
                f"Config atualizado: {alteracao.caminho} "
                f"({alteracao.ocorrencias} ocorrencia(s), backup em {backup.name})"
            )
    return backups


def restaurar_backups(arquivos_backup: Sequence[tuple[Path, Path]]) -> None:
    for destino, backup in reversed(list(arquivos_backup)):
        shutil.copy2(backup, destino)


def caminho_mapeamento_base(pasta_interface: Path) -> Path:
    return pasta_interface / NOME_ARQUIVO_MAPEAMENTO_BASE


def montar_linha_mapeamento_base(nome_original: str, nome_isolado: str) -> str:
    carimbo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"{carimbo} | {nome_original} --> {nome_isolado}"


def registrar_mapeamento_base(
    pasta_interface: Path,
    nome_original: str,
    nome_isolado: str,
    logger: Callable[[str], None] | None = None,
) -> str:
    arquivo_mapeamento = caminho_mapeamento_base(pasta_interface)
    linha = montar_linha_mapeamento_base(nome_original, nome_isolado)
    with arquivo_mapeamento.open("a", encoding="utf-8", newline="\n") as arquivo:
        arquivo.write(f"{linha}\n")
    logar(logger, f"Mapeamento registrado em {arquivo_mapeamento.name}: {nome_original} --> {nome_isolado}")
    return linha


def remover_ultimo_mapeamento_registrado(pasta_interface: Path, linha_registrada: str) -> None:
    arquivo_mapeamento = caminho_mapeamento_base(pasta_interface)
    if not arquivo_mapeamento.exists():
        return

    linhas = arquivo_mapeamento.read_text(encoding="utf-8").splitlines()
    if not linhas or linhas[-1].strip() != linha_registrada.strip():
        return

    conteudo = "\n".join(linhas[:-1])
    if conteudo:
        conteudo += "\n"
    arquivo_mapeamento.write_text(conteudo, encoding="utf-8", newline="\n")


def extrair_mapeamento_base_linha(linha: str) -> tuple[str, str] | None:
    texto = linha.strip()
    if not texto:
        return None

    if "|" in texto:
        texto = texto.split("|", 1)[1].strip()

    if "-->" not in texto:
        return None

    nome_original, nome_isolado = [parte.strip() for parte in texto.split("-->", 1)]
    if not nome_original or not nome_isolado:
        return None
    return nome_original, nome_isolado


def ler_ultimo_nome_original_mapeado(
    pasta_interface: Path,
    nome_isolado: str = NOME_BASE_ISOLADA,
) -> str | None:
    arquivo_mapeamento = caminho_mapeamento_base(pasta_interface)
    if not arquivo_mapeamento.exists():
        return None

    for linha in reversed(arquivo_mapeamento.read_text(encoding="utf-8").splitlines()):
        mapeamento = extrair_mapeamento_base_linha(linha)
        if mapeamento is None:
            continue

        nome_original_linha, nome_isolado_linha = mapeamento
        if nome_isolado_linha.lower() == nome_isolado.lower():
            return nome_original_linha

    return None


def preparar_base_para_atualizacao(
    pasta_interface: Path,
    pasta_setups: Path,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> ResultadoAnalise:
    verificar_cancelamento(cancelar_evento, "preparo da base")
    analise = analisar_atualizacao(pasta_interface, pasta_setups)
    banco_atual = analise.banco_principal.caminho
    banco_isolado = banco_atual.with_name(NOME_BASE_ISOLADA)

    if analise.base_isolada:
        if logger is not None:
            logger("A base ja esta isolada com o nome usado para atualizacao.")
        return analise

    if banco_isolado.exists():
        raise RuntimeError(
            f"Ja existe uma base com o nome de isolamento em {banco_isolado}. "
            "Renomeie ou remova esse arquivo antes de continuar."
        )

    alteracoes = planejar_alteracoes_configuracao(pasta_interface, banco_atual.name, NOME_BASE_ISOLADA)
    if not alteracoes:
        raise RuntimeError(
            f"Nenhum arquivo .ini referencia {banco_atual.name}. "
            "A troca do nome da base foi abortada para evitar deixar a maquina sem acesso."
        )

    backups: list[tuple[Path, Path]] = []
    linha_mapeamento: str | None = None

    try:
        linha_mapeamento = registrar_mapeamento_base(
            pasta_interface,
            banco_atual.name,
            banco_isolado.name,
            logger=logger,
        )
        banco_atual.rename(banco_isolado)
        if logger is not None:
            logger(f"Banco renomeado: {banco_atual.name} -> {banco_isolado.name}")

        backups = aplicar_alteracoes_configuracao(alteracoes, logger=logger)

        analise_final = analisar_atualizacao(pasta_interface, pasta_setups)
        if logger is not None:
            logger("Validacao concluida: a base isolada continuou acessivel pela configuracao da maquina.")
        return analise_final
    except Exception:
        if backups:
            restaurar_backups(backups)
        if banco_isolado.exists() and not banco_atual.exists():
            banco_isolado.rename(banco_atual)
        if linha_mapeamento is not None:
            remover_ultimo_mapeamento_registrado(pasta_interface, linha_mapeamento)
        raise


def restaurar_nome_original_base(
    pasta_interface: Path,
    pasta_setups: Path,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> ResultadoAnalise:
    verificar_cancelamento(cancelar_evento, "restauracao da base")
    analise = analisar_atualizacao(pasta_interface, pasta_setups)
    banco_atual = analise.banco_principal.caminho
    nome_original_registrado = ler_ultimo_nome_original_mapeado(pasta_interface)

    if nome_original_registrado and banco_atual.name.lower() == nome_original_registrado.lower():
        if logger is not None:
            logger(f"A base ja esta com o nome original registrado: {nome_original_registrado}.")
        return analise

    if banco_atual.name.lower() == NOME_BASE_ORIGINAL.lower():
        if logger is not None:
            logger("A base ja esta com o nome original padrao.")
        return analise

    if banco_atual.name.lower() != NOME_BASE_ISOLADA.lower():
        raise RuntimeError(
            f"A restauracao automatica espera a base com o nome {NOME_BASE_ISOLADA}, "
            f"mas a base atual e {banco_atual.name}."
        )

    nome_original_destino = nome_original_registrado or NOME_BASE_ORIGINAL
    if logger is not None and nome_original_registrado is None:
        logger(
            f"Mapa de nomes nao encontrado em {caminho_mapeamento_base(pasta_interface).name}; "
            f"sera usado o nome padrao {NOME_BASE_ORIGINAL}."
        )

    banco_original = banco_atual.with_name(nome_original_destino)
    if banco_original.exists():
        raise RuntimeError(f"Ja existe um arquivo com o nome original em {banco_original}.")

    alteracoes = planejar_alteracoes_configuracao(pasta_interface, banco_atual.name, nome_original_destino)
    backups: list[tuple[Path, Path]] = []

    try:
        banco_atual.rename(banco_original)
        if logger is not None:
            logger(f"Banco restaurado: {banco_atual.name} -> {banco_original.name}")

        backups = aplicar_alteracoes_configuracao(alteracoes, logger=logger)
        return analisar_atualizacao(pasta_interface, pasta_setups)
    except Exception:
        if backups:
            restaurar_backups(backups)
        if banco_original.exists() and not banco_atual.exists():
            banco_original.rename(banco_atual)
        raise


def formatar_resumo_analise(analise: ResultadoAnalise) -> str:
    linhas = [
        f"Pasta analisada: {analise.pasta_interface}",
        f"Banco principal: {analise.banco_principal.caminho}",
        f"Origem: {analise.banco_principal.origem}",
        f"Tabela encontrada: {analise.tabela_compilacao}",
        f"{CAMPO_COMPILACAO}: {formatar_versao(analise.versao_atual)}",
        f"Base isolada: {'SIM' if analise.base_isolada else 'NAO'}",
        "Fila de atualizacao:",
    ]

    if analise.fila_atualizacao:
        linhas.extend(f"- {setup.descricao}" for setup in analise.fila_atualizacao)
    else:
        linhas.append(f"- {CHAVE_FILA_VAZIA}")

    return "\n".join(linhas)


def montar_argumentos_instalador(argumentos_livres: str) -> list[str]:
    if not argumentos_livres.strip():
        return []
    return shlex.split(argumentos_livres, posix=False)


def executar_atualizacoes_mensais(
    pasta_interface: Path,
    pasta_setups: Path,
    argumentos_instalador: str = "",
    preparar_antes: bool = True,
    logger: Callable[[str], None] | None = None,
    cancelar_evento: threading.Event | None = None,
) -> ResultadoAnalise:
    if preparar_antes:
        analise = preparar_base_para_atualizacao(
            pasta_interface,
            pasta_setups,
            logger=logger,
            cancelar_evento=cancelar_evento,
        )
    else:
        analise = analisar_atualizacao(pasta_interface, pasta_setups)

    argumentos = montar_argumentos_instalador(argumentos_instalador)

    total_executado = 0
    while True:
        verificar_cancelamento(cancelar_evento, "execucao da fila de atualizacao")
        fila_pendente = analise.fila_atualizacao
        if not fila_pendente:
            if total_executado == 0 and logger is not None:
                logger("Nao ha atualizacoes pendentes para executar.")
            break

        setup = fila_pendente[0]
        total_executado += 1

        if logger is not None:
            logger(
                f"Abrindo setup {total_executado}: {setup.caminho.name} "
                f"(restantes apos este: {len(fila_pendente) - 1})"
            )

        automatizar_setup(
            setup,
            pasta_interface,
            argumentos,
            logger=logger,
            cancelar_evento=cancelar_evento,
        )

        analise = analisar_atualizacao(pasta_interface, pasta_setups)
        if logger is not None:
            logger(
                "Setup concluido. "
                f"Versao atual detectada no banco: {formatar_versao(analise.versao_atual)}"
            )

    if total_executado > 0:
        if logger is not None:
            logger("Todas as atualizacoes foram concluidas. Restaurando o nome original da base.")
        analise = restaurar_nome_original_base(
            pasta_interface,
            pasta_setups,
            logger=logger,
            cancelar_evento=cancelar_evento,
        )

        aplicar_correcao_grid_localizacao_produtos(analise, logger=logger)
        analise = analisar_atualizacao(pasta_interface, pasta_setups)

    return analise


class SeletorPastaDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        titulo: str,
        pasta_inicial: Path,
        cores: dict[str, str],
        fontes: dict[str, tkfont.Font],
    ) -> None:
        super().__init__(master)
        self.resultado = ""
        self.cores = cores
        self.fontes = fontes
        self.pasta_atual = pasta_inicial
        self.unidades = self._listar_unidades()
        self.caminho_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Selecione uma pasta na arvore ou digite o caminho completo.")
        self._nos_carregados: set[str] = set()

        self.title(titulo)
        self.withdraw()
        aplicar_geometria_inicial(self, 760, 520, 640, 440)
        self.configure(bg=self.cores["bg"])
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

        self._montar_layout()
        self._centralizar(master)
        self._popular_raizes()
        self._abrir_pasta(pasta_inicial)

    def mostrar(self) -> str:
        self.deiconify()
        self.transient(self.master)
        self.lift(self.master)
        self.attributes("-topmost", True)
        self.after(150, self._desativar_topmost)
        self.wait_visibility()
        self.grab_set()
        self.focus_force()
        self.entrada_caminho.focus_set()
        self.wait_window()
        return self.resultado

    def _desativar_topmost(self) -> None:
        try:
            if self.winfo_exists():
                self.attributes("-topmost", False)
        except tk.TclError:
            pass

    def _montar_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        conteudo = tk.Frame(
            self,
            bg=self.cores["surface"],
            highlightbackground=self.cores["border"],
            highlightthickness=1,
            padx=18,
            pady=18,
        )
        conteudo.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        conteudo.grid_columnconfigure(0, weight=1)
        conteudo.grid_rowconfigure(3, weight=1)

        cabecalho = tk.Frame(
            conteudo,
            bg=self.cores["surface_alt"],
            highlightbackground=self.cores["border_strong"],
            highlightthickness=1,
            padx=18,
            pady=16,
        )
        cabecalho.grid(row=0, column=0, sticky="ew")
        cabecalho.grid_columnconfigure(0, weight=1)

        tk.Label(
            cabecalho,
            text="Localizar pasta",
            bg=self.cores["surface_alt"],
            fg=self.cores["primary"],
            font=self.fontes["secao"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            cabecalho,
            text="Use a unidade, navegue por subpastas ou digite o caminho completo.",
            bg=self.cores["surface_alt"],
            fg=self.cores["muted"],
            font=self.fontes["subtitulo"],
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        tk.Label(
            cabecalho,
            text="Duplo clique ou Enter para confirmar",
            bg=self.cores["highlight"],
            fg=self.cores["primary"],
            font=self.fontes["badge"],
            padx=12,
            pady=7,
        ).grid(row=0, column=1, rowspan=2, sticky="ne", padx=(14, 0))

        barra = tk.Frame(conteudo, bg=self.cores["surface"])
        barra.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        barra.grid_columnconfigure(0, weight=1)

        self.entrada_caminho = ttk.Entry(barra, textvariable=self.caminho_var, style="Path.TEntry")
        self.entrada_caminho.grid(row=0, column=0, sticky="ew")
        self.entrada_caminho.bind("<Return>", self._abrir_caminho_digitado)

        ttk.Button(
            barra,
            text="Ir",
            style="Soft.TButton",
            command=self._abrir_caminho_digitado,
        ).grid(row=0, column=1, sticky="e", padx=(10, 0))

        ttk.Button(
            barra,
            text="Subir",
            style="Soft.TButton",
            command=self._subir_nivel,
        ).grid(row=0, column=2, sticky="e", padx=(10, 0))

        lista_frame = tk.Frame(
            conteudo,
            bg=self.cores["surface_soft"],
            highlightbackground=self.cores["border"],
            highlightthickness=1,
            padx=10,
            pady=10,
        )
        lista_frame.grid(row=3, column=0, sticky="nsew", pady=(14, 0))
        lista_frame.grid_columnconfigure(0, weight=1)
        lista_frame.grid_rowconfigure(0, weight=1)

        self.arvore_pastas = ttk.Treeview(
            lista_frame,
            show="tree",
            selectmode="browse",
            style="App.Treeview",
        )
        self.arvore_pastas.grid(row=0, column=0, sticky="nsew")
        self.arvore_pastas.bind("<<TreeviewOpen>>", self._ao_expandir_no)
        self.arvore_pastas.bind("<<TreeviewSelect>>", self._ao_selecionar_no)
        self.arvore_pastas.bind("<Double-Button-1>", self._ao_duplo_clique)
        self.arvore_pastas.bind("<Return>", self._confirmar_selecao)

        scrollbar = ttk.Scrollbar(
            lista_frame,
            orient="vertical",
            command=self.arvore_pastas.yview,
            style="App.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.arvore_pastas.configure(yscrollcommand=scrollbar.set)

        rodape = tk.Frame(conteudo, bg=self.cores["surface"])
        rodape.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        rodape.grid_columnconfigure(0, weight=1)

        tk.Label(
            rodape,
            textvariable=self.status_var,
            bg=self.cores["surface"],
            fg=self.cores["muted"],
            font=self.fontes["subtitulo"],
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, sticky="w")

        botoes = tk.Frame(rodape, bg=self.cores["surface"])
        botoes.grid(row=0, column=1, sticky="e", padx=(12, 0))

        ttk.Button(
            botoes,
            text="Usar selecionada",
            style="Soft.TButton",
            command=self._confirmar_selecao,
        ).grid(row=0, column=0, padx=(0, 8))

        ttk.Button(
            botoes,
            text="Usar pasta atual",
            style="Primary.TButton",
            command=self._selecionar_pasta_atual,
        ).grid(row=0, column=1, padx=(0, 8))

        ttk.Button(
            botoes,
            text="Cancelar",
            style="Soft.TButton",
            command=self._cancelar,
        ).grid(row=0, column=2)

    def _centralizar(self, master: tk.Misc) -> None:
        self.update_idletasks()
        largura = self.winfo_width()
        altura = self.winfo_height()
        largura_tela = max(1, self.winfo_screenwidth())
        altura_tela = max(1, self.winfo_screenheight())
        x = master.winfo_rootx() + max((master.winfo_width() - largura) // 2, 20)
        y = master.winfo_rooty() + max((master.winfo_height() - altura) // 2, 20)
        x = max(0, min(x, largura_tela - largura))
        y = max(0, min(y, altura_tela - altura))
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    def _listar_unidades(self) -> list[str]:
        try:
            bruto = win32api.GetLogicalDriveStrings()
        except Exception:  # noqa: BLE001
            bruto = ""

        unidades = [
            unidade
            for unidade in bruto.split("\x00")
            if unidade and Path(unidade).exists()
        ]
        if not unidades:
            unidades = [str(PASTA_APLICACAO.anchor or Path.home().anchor or Path.home())]
        return unidades

    def _popular_raizes(self) -> None:
        for unidade in self.unidades:
            caminho = self._normalizar_caminho(unidade)
            iid = self._iid_para_caminho(caminho)
            if self.arvore_pastas.exists(iid):
                continue
            self.arvore_pastas.insert("", "end", iid=iid, text=str(caminho), open=False)
            self._garantir_placeholder(iid)

    def _normalizar_caminho(self, caminho: str | Path) -> Path:
        texto = str(caminho).strip().strip('"')
        if re.fullmatch(r"[a-zA-Z]:", texto):
            texto = f"{texto}\\"
        return Path(texto).expanduser()

    def _iid_para_caminho(self, caminho: Path) -> str:
        return str(self._normalizar_caminho(caminho))

    def _placeholder_iid(self, iid_pai: str) -> str:
        return f"__placeholder__::{iid_pai}"

    def _garantir_placeholder(self, iid_pai: str) -> None:
        if iid_pai and not self.arvore_pastas.exists(iid_pai):
            return
        if self.arvore_pastas.exists(self._placeholder_iid(iid_pai)):
            return
        self.arvore_pastas.insert(iid_pai, "end", iid=self._placeholder_iid(iid_pai), text="...")

    def _carregar_subpastas(self, caminho: Path) -> None:
        iid = self._iid_para_caminho(caminho)
        if iid in self._nos_carregados:
            return

        placeholder = self._placeholder_iid(iid)
        if self.arvore_pastas.exists(placeholder):
            self.arvore_pastas.delete(placeholder)

        try:
            subpastas = sorted(
                (item for item in caminho.iterdir() if item.is_dir()),
                key=lambda item: item.name.lower(),
            )
        except OSError as exc:
            self.status_var.set(f"Nao foi possivel listar '{caminho}': {exc}")
            return

        for pasta in subpastas:
            iid_filho = self._iid_para_caminho(pasta)
            if self.arvore_pastas.exists(iid_filho):
                continue
            self.arvore_pastas.insert(iid, "end", iid=iid_filho, text=pasta.name, open=False)
            self._garantir_placeholder(iid_filho)

        self._nos_carregados.add(iid)

    def _abrir_caminho_digitado(self, _evento: object = None) -> None:
        caminho = self.caminho_var.get().strip()
        if not caminho:
            return
        self._abrir_pasta(self._normalizar_caminho(caminho))

    def _subir_nivel(self) -> None:
        pai = self.pasta_atual.parent
        if pai == self.pasta_atual:
            self.status_var.set("Voce ja esta na raiz desta unidade.")
            return
        self._abrir_pasta(pai)

    def _ao_expandir_no(self, _evento: object = None) -> None:
        pasta = self._obter_pasta_em_foco()
        if pasta is None:
            return
        self._carregar_subpastas(pasta)

    def _ao_selecionar_no(self, _evento: object = None) -> None:
        pasta = self._obter_pasta_selecionada()
        if pasta is None:
            return
        self.pasta_atual = pasta
        self.caminho_var.set(str(pasta))
        self.status_var.set("Pasta selecionada. Clique em 'Usar selecionada' para confirmar.")

    def _ao_duplo_clique(self, _evento: object = None) -> None:
        pasta = self._obter_pasta_selecionada()
        if pasta is None:
            return
        self._abrir_pasta(pasta)

    def _selecionar_pasta_atual(self) -> None:
        self.resultado = str(self.pasta_atual)
        self.destroy()

    def _cancelar(self) -> None:
        self.resultado = ""
        self.destroy()

    def _obter_pasta_em_foco(self) -> Path | None:
        iid = self.arvore_pastas.focus()
        if not iid or iid.startswith("__placeholder__::"):
            return None
        return self._normalizar_caminho(iid)

    def _obter_pasta_selecionada(self) -> Path | None:
        selecao = self.arvore_pastas.selection()
        if not selecao:
            self.status_var.set("Selecione uma pasta na arvore ou use a pasta atual.")
            return None

        iid = selecao[0]
        if iid.startswith("__placeholder__::"):
            return None
        return self._normalizar_caminho(iid)

    def _confirmar_selecao(self, _evento: object = None) -> None:
        pasta = self._obter_pasta_selecionada()
        if pasta is None:
            self._selecionar_pasta_atual()
            return
        self.resultado = str(pasta)
        self.destroy()

    def _expandir_ate_caminho(self, caminho: Path) -> None:
        linhagem: list[Path] = []
        atual = self._normalizar_caminho(caminho)
        while True:
            linhagem.append(atual)
            if atual.parent == atual:
                break
            atual = atual.parent
        linhagem.reverse()

        for indice, parte in enumerate(linhagem):
            iid = self._iid_para_caminho(parte)
            if not self.arvore_pastas.exists(iid):
                if indice == 0:
                    self.arvore_pastas.insert("", "end", iid=iid, text=str(parte), open=False)
                    self._garantir_placeholder(iid)
                else:
                    pai = self._iid_para_caminho(linhagem[indice - 1])
                    self._carregar_subpastas(linhagem[indice - 1])
                    if not self.arvore_pastas.exists(iid):
                        self.arvore_pastas.insert(pai, "end", iid=iid, text=parte.name or str(parte), open=False)
                        self._garantir_placeholder(iid)

            if indice < len(linhagem) - 1:
                self._carregar_subpastas(parte)
                self.arvore_pastas.item(iid, open=True)

        destino_iid = self._iid_para_caminho(caminho)
        if self.arvore_pastas.exists(destino_iid):
            self.arvore_pastas.selection_set(destino_iid)
            self.arvore_pastas.focus(destino_iid)
            self.arvore_pastas.see(destino_iid)

    def _abrir_pasta(self, caminho: Path) -> None:
        caminho = self._normalizar_caminho(caminho)

        try:
            if not caminho.exists():
                raise FileNotFoundError(f"A pasta '{caminho}' nao existe.")
            if not caminho.is_dir():
                raise NotADirectoryError(f"'{caminho}' nao e uma pasta.")
        except OSError as exc:
            messagebox.showerror(APP_TITULO, str(exc), parent=self)
            self.status_var.set("Nao foi possivel abrir o caminho informado.")
            return

        self.pasta_atual = caminho
        self.caminho_var.set(str(caminho))
        self._expandir_ate_caminho(caminho)
        self._carregar_subpastas(caminho)
        self.status_var.set("Pasta carregada. Navegue pela arvore ou confirme a pasta atual.")


class AtualizadorInterfaceApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITULO)
        aplicar_geometria_inicial(self, 1120, 820, 980, 700)
        self._icone_imagem: tk.PhotoImage | None = None

        self.fila_eventos: queue.Queue[tuple[str, object]] = queue.Queue()
        self.trabalho_ativo = False
        self.analise_atual: ResultadoAnalise | None = None
        self.cancelamento_evento = threading.Event()
        self.cores = {
            "bg": "#f4ecdf",
            "surface": "#fffaf2",
            "surface_soft": "#fbf3e6",
            "surface_alt": "#efdfc7",
            "border": "#dbc3a0",
            "border_strong": "#b8935f",
            "ink": "#1f2b2d",
            "muted": "#6a7067",
            "primary": "#163f4a",
            "primary_active": "#0d2e36",
            "secondary": "#c7742d",
            "secondary_active": "#a95c1f",
            "tertiary": "#647e56",
            "tertiary_soft": "#e2ebdd",
            "hero": "#17353d",
            "hero_secondary": "#2a5560",
            "highlight": "#f2d5a2",
            "disabled": "#bec6c8",
            "log_bg": "#10212c",
            "log_bg_soft": "#163140",
            "log_fg": "#f5efe3",
            "ok": "#3c7a53",
        }

        self.pasta_interface_var = tk.StringVar(value=str(PASTA_PADRAO))
        self.pasta_setups_var = tk.StringVar(value=str(PASTA_SETUPS_PADRAO))
        self.banco_var = tk.StringVar(value="-")
        self.versao_var = tk.StringVar(value="-")
        self.isolada_var = tk.StringVar(value="-")
        self.tabela_var = tk.StringVar(value="-")
        self.argumentos_var = tk.StringVar(value="")
        self.fila_resumo_var = tk.StringVar(value="Nenhuma fila carregada")
        self.status_execucao_var = tk.StringVar(value="Pronto para analisar a base.")

        self._configurar_estilo()
        self._configurar_icone()
        self._montar_layout()
        self.after(150, self._processar_fila_eventos)
        self.after(250, self.analisar)

    def _configurar_icone(self) -> None:
        try:
            if ARQUIVO_ICONE.exists():
                self.iconbitmap(default=str(ARQUIVO_ICONE))
        except tk.TclError:
            pass

        try:
            if ARQUIVO_LOGO.exists():
                self._icone_imagem = tk.PhotoImage(file=str(ARQUIVO_LOGO))
                self.iconphoto(True, self._icone_imagem)
        except tk.TclError:
            self._icone_imagem = None

    def _configurar_estilo(self) -> None:
        self.option_add("*tearOff", False)
        self.configure(bg=self.cores["bg"])

        fonte_base = primeira_fonte_disponivel(
            ("Aptos", "Segoe UI", "Verdana", "Tahoma"),
            "Segoe UI",
        )
        fonte_destaque = primeira_fonte_disponivel(
            ("Bahnschrift SemiBold", "Bahnschrift", "Franklin Gothic Demi", "Trebuchet MS Bold"),
            "Segoe UI Semibold",
        )
        fonte_texto = primeira_fonte_disponivel(
            ("Segoe UI", "Aptos", "Verdana", "Tahoma"),
            "Segoe UI",
        )
        fonte_mono = primeira_fonte_disponivel(
            ("Cascadia Mono", "Consolas", "Lucida Console", "Courier New"),
            "Consolas",
        )

        fonte_padrao = tkfont.nametofont("TkDefaultFont")
        fonte_padrao.configure(family=fonte_texto, size=10)
        tkfont.nametofont("TkMenuFont").configure(family=fonte_texto, size=10)
        tkfont.nametofont("TkTextFont").configure(family=fonte_mono, size=10)

        self.fontes = {
            "titulo": tkfont.Font(family=fonte_destaque, size=26, weight="bold"),
            "hero_titulo": tkfont.Font(family=fonte_destaque, size=30, weight="bold"),
            "hero_kicker": tkfont.Font(family=fonte_base, size=10, weight="bold"),
            "hero_texto": tkfont.Font(family=fonte_texto, size=11),
            "subtitulo": tkfont.Font(family=fonte_texto, size=10),
            "secao": tkfont.Font(family=fonte_destaque, size=13, weight="bold"),
            "secao_suave": tkfont.Font(family=fonte_base, size=9, weight="bold"),
            "card_titulo": tkfont.Font(family=fonte_base, size=9, weight="bold"),
            "card_valor": tkfont.Font(family=fonte_destaque, size=13, weight="bold"),
            "card_destaque": tkfont.Font(family=fonte_destaque, size=18, weight="bold"),
            "botao": tkfont.Font(family=fonte_base, size=10, weight="bold"),
            "rodape": tkfont.Font(family=fonte_texto, size=9),
            "badge": tkfont.Font(family=fonte_base, size=9, weight="bold"),
            "mono": tkfont.Font(family=fonte_mono, size=10),
        }

        self.estilo = ttk.Style(self)
        try:
            self.estilo.theme_use("clam")
        except tk.TclError:
            pass

        self.estilo.configure("App.TFrame", background=self.cores["bg"])
        self.estilo.configure(
            "Path.TEntry",
            padding=(12, 10),
            fieldbackground=self.cores["surface_soft"],
            foreground=self.cores["ink"],
            bordercolor=self.cores["border"],
            lightcolor=self.cores["surface_soft"],
            darkcolor=self.cores["border"],
            insertcolor=self.cores["ink"],
            relief="flat",
        )
        self.estilo.configure(
            "Primary.TButton",
            padding=(16, 11),
            font=self.fontes["botao"],
            background=self.cores["primary"],
            foreground="#ffffff",
            borderwidth=0,
        )
        self.estilo.map(
            "Primary.TButton",
            background=[
                ("active", self.cores["primary_active"]),
                ("disabled", self.cores["disabled"]),
            ],
            foreground=[("disabled", "#eff3f5")],
        )
        self.estilo.configure(
            "Accent.TButton",
            padding=(16, 11),
            font=self.fontes["botao"],
            background=self.cores["secondary"],
            foreground="#ffffff",
            borderwidth=0,
        )
        self.estilo.map(
            "Accent.TButton",
            background=[
                ("active", self.cores["secondary_active"]),
                ("disabled", self.cores["disabled"]),
            ],
            foreground=[("disabled", "#eff3f5")],
        )
        self.estilo.configure(
            "Soft.TButton",
            padding=(14, 10),
            font=self.fontes["botao"],
            background=self.cores["surface_alt"],
            foreground=self.cores["primary"],
            borderwidth=1,
            bordercolor=self.cores["border"],
        )
        self.estilo.map(
            "Soft.TButton",
            background=[
                ("active", self.cores["highlight"]),
                ("disabled", self.cores["disabled"]),
            ],
            foreground=[
                ("active", self.cores["primary"]),
                ("disabled", "#f7f7f7"),
            ],
        )
        self.estilo.configure(
            "App.Treeview",
            background=self.cores["surface_soft"],
            fieldbackground=self.cores["surface_soft"],
            foreground=self.cores["ink"],
            bordercolor=self.cores["border"],
            rowheight=28,
            font=self.fontes["subtitulo"],
            relief="flat",
        )
        self.estilo.map(
            "App.Treeview",
            background=[("selected", self.cores["primary"])],
            foreground=[("selected", "#ffffff")],
        )
        self.estilo.configure(
            "App.Treeview.Heading",
            background=self.cores["surface_alt"],
            foreground=self.cores["primary"],
            font=self.fontes["card_titulo"],
            padding=(10, 8),
            relief="flat",
        )
        self.estilo.map(
            "App.Treeview.Heading",
            background=[("active", self.cores["highlight"])],
            foreground=[("active", self.cores["primary"])],
        )
        self.estilo.configure(
            "App.Vertical.TScrollbar",
            troughcolor=self.cores["surface_soft"],
            background=self.cores["surface_alt"],
            bordercolor=self.cores["border"],
            arrowcolor=self.cores["primary"],
            darkcolor=self.cores["surface_alt"],
            lightcolor=self.cores["surface_alt"],
            arrowsize=14,
        )
        self.estilo.configure(
            "App.Horizontal.TProgressbar",
            troughcolor=self.cores["surface_soft"],
            bordercolor=self.cores["surface_soft"],
            background=self.cores["secondary"],
            lightcolor=self.cores["secondary"],
            darkcolor=self.cores["secondary"],
        )

    def _criar_superficie(
        self,
        pai: tk.Misc,
        padx: int = 18,
        pady: int = 18,
        background: str | None = None,
        border: str | None = None,
    ) -> tk.Frame:
        fundo = background or self.cores["surface"]
        borda = border or self.cores["border"]
        frame = tk.Frame(
            pai,
            bg=fundo,
            highlightbackground=borda,
            highlightthickness=1,
            bd=0,
            padx=padx,
            pady=pady,
        )
        return frame

    def _criar_selo(
        self,
        pai: tk.Misc,
        texto: str,
        background: str | None = None,
        foreground: str | None = None,
        font_key: str = "badge",
        padx: int = 12,
        pady: int = 7,
    ) -> tk.Label:
        return tk.Label(
            pai,
            text=texto,
            bg=background or self.cores["surface_alt"],
            fg=foreground or self.cores["primary"],
            font=self.fontes[font_key],
            padx=padx,
            pady=pady,
        )

    def _criar_titulo_secao(self, pai: tk.Misc, titulo: str, descricao: str | None = None) -> tk.Frame:
        cabecalho = tk.Frame(pai, bg=pai.cget("bg"))
        tk.Label(
            cabecalho,
            text=titulo,
            bg=pai.cget("bg"),
            fg=self.cores["ink"],
            font=self.fontes["secao"],
            anchor="w",
        ).pack(anchor="w")

        if descricao:
            tk.Label(
                cabecalho,
                text=descricao,
                bg=pai.cget("bg"),
                fg=self.cores["muted"],
                font=self.fontes["subtitulo"],
                anchor="w",
                justify="left",
            ).pack(anchor="w", pady=(4, 0))

        return cabecalho

    def _criar_cartao_status(self, pai: tk.Misc, titulo: str, variavel: tk.StringVar, destaque: bool = False) -> None:
        fundo = self.cores["surface_soft"] if destaque else self.cores["surface"]
        cartao = self._criar_superficie(
            pai,
            padx=16,
            pady=14,
            background=fundo,
            border=self.cores["border_strong"] if destaque else self.cores["border"],
        )
        tk.Frame(
            cartao,
            bg=self.cores["secondary"] if destaque else self.cores["tertiary"],
            height=4,
        ).pack(fill="x")

        corpo = tk.Frame(cartao, bg=fundo)
        corpo.pack(fill="both", expand=True, pady=(12, 0))
        tk.Label(
            corpo,
            text=titulo.upper(),
            bg=fundo,
            fg=self.cores["muted"],
            font=self.fontes["card_titulo"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            corpo,
            textvariable=variavel,
            bg=fundo,
            fg=self.cores["primary"] if destaque else self.cores["ink"],
            font=self.fontes["card_destaque"] if destaque else self.fontes["card_valor"],
            justify="left",
            wraplength=360,
            anchor="w",
        ).pack(anchor="w", pady=(8, 0), fill="x")
        cartao.pack(side="left", fill="both", expand=True)

    def _atualizar_scrollregion_principal(self, _evento: tk.Event | None = None) -> None:
        self.canvas_principal.configure(scrollregion=self.canvas_principal.bbox("all"))

    def _ajustar_largura_container_principal(self, evento: tk.Event) -> None:
        self.canvas_principal.itemconfigure(self.container_canvas_id, width=evento.width)

    def _rolar_principal_mousewheel(self, evento: tk.Event) -> str | None:
        classe = ""
        try:
            classe = evento.widget.winfo_class()
        except Exception:  # noqa: BLE001
            pass

        if classe in {"Text", "Listbox"}:
            return None

        delta = 0
        if getattr(evento, "delta", 0):
            delta = int(-evento.delta / 120)
        elif getattr(evento, "num", None) == 4:
            delta = -1
        elif getattr(evento, "num", None) == 5:
            delta = 1

        if delta:
            self.canvas_principal.yview_scroll(delta, "units")
            return "break"
        return None

    def _montar_layout(self) -> None:
        self.frame_rolagem = tk.Frame(self, bg=self.cores["bg"])
        self.frame_rolagem.pack(fill="both", expand=True)
        self.frame_rolagem.grid_columnconfigure(0, weight=1)
        self.frame_rolagem.grid_rowconfigure(0, weight=1)

        self.canvas_principal = tk.Canvas(
            self.frame_rolagem,
            bg=self.cores["bg"],
            highlightthickness=0,
            bd=0,
        )
        self.canvas_principal.grid(row=0, column=0, sticky="nsew")

        self.barra_rolagem_principal = ttk.Scrollbar(
            self.frame_rolagem,
            orient="vertical",
            command=self.canvas_principal.yview,
            style="App.Vertical.TScrollbar",
        )
        self.barra_rolagem_principal.grid(row=0, column=1, sticky="ns")
        self.canvas_principal.configure(yscrollcommand=self.barra_rolagem_principal.set)

        self.container = tk.Frame(self.canvas_principal, bg=self.cores["bg"], padx=28, pady=24)
        self.container_canvas_id = self.canvas_principal.create_window((0, 0), window=self.container, anchor="nw")
        self.container.bind("<Configure>", self._atualizar_scrollregion_principal)
        self.canvas_principal.bind("<Configure>", self._ajustar_largura_container_principal)
        self.bind_all("<MouseWheel>", self._rolar_principal_mousewheel, add="+")
        self.bind_all("<Button-4>", self._rolar_principal_mousewheel, add="+")
        self.bind_all("<Button-5>", self._rolar_principal_mousewheel, add="+")

        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(5, weight=1)

        hero = tk.Frame(
            self.container,
            bg=self.cores["hero"],
            highlightbackground=self.cores["hero_secondary"],
            highlightthickness=1,
            padx=28,
            pady=24,
        )
        hero.grid(row=0, column=0, sticky="ew")
        hero.grid_columnconfigure(0, weight=3)
        hero.grid_columnconfigure(1, weight=2)

        hero_esquerda = tk.Frame(hero, bg=self.cores["hero"])
        hero_esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, 18))

        self._criar_selo(
            hero_esquerda,
            "PAINEL DE ATUALIZACAO",
            background=self.cores["highlight"],
            foreground=self.cores["primary"],
        ).pack(anchor="w")

        tk.Label(
            hero_esquerda,
            text=APP_TITULO,
            bg=self.cores["hero"],
            fg="#fff8ed",
            font=self.fontes["hero_titulo"],
            anchor="w",
        ).pack(anchor="w", pady=(16, 0))

        tk.Label(
            hero_esquerda,
            text=(
                "Atualize a base com mais clareza, acompanhe cada etapa em tempo real "
                "e mantenha o processo seguro do primeiro setup ao fechamento final."
            ),
            bg=self.cores["hero"],
            fg="#d7e4e0",
            font=self.fontes["hero_texto"],
            anchor="w",
            justify="left",
            wraplength=580,
        ).pack(anchor="w", pady=(10, 0))

        chips_hero = tk.Frame(hero_esquerda, bg=self.cores["hero"])
        chips_hero.pack(anchor="w", pady=(18, 0))
        for texto in ("Leitura segura", "Fila mensal", "Restauracao automatica"):
            self._criar_selo(
                chips_hero,
                texto,
                background=self.cores["hero_secondary"],
                foreground="#fff8ed",
                padx=11,
                pady=6,
            ).pack(side="left", padx=(0, 8))

        painel_resumo = tk.Frame(
            hero,
            bg=self.cores["surface"],
            highlightbackground=self.cores["border_strong"],
            highlightthickness=1,
            padx=18,
            pady=18,
        )
        painel_resumo.grid(row=0, column=1, sticky="nsew")
        painel_resumo.grid_columnconfigure(0, weight=1)

        tk.Label(
            painel_resumo,
            text="Visao rapida",
            bg=self.cores["surface"],
            fg=self.cores["primary"],
            font=self.fontes["secao"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        resumo_metricas = tk.Frame(painel_resumo, bg=self.cores["surface"])
        resumo_metricas.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        resumo_metricas.grid_columnconfigure(0, weight=1)
        resumo_metricas.grid_columnconfigure(1, weight=1)

        for indice, (titulo, variavel) in enumerate(
            (
                ("Fila prevista", self.fila_resumo_var),
                ("Versao lida", self.versao_var),
                ("Base isolada", self.isolada_var),
                ("Tabela", self.tabela_var),
            )
        ):
            linha = indice // 2
            coluna = indice % 2
            bloco = tk.Frame(resumo_metricas, bg=self.cores["surface"])
            bloco.grid(row=linha, column=coluna, sticky="ew", padx=(0, 12 if coluna == 0 else 0), pady=(0, 12))
            tk.Label(
                bloco,
                text=titulo.upper(),
                bg=self.cores["surface"],
                fg=self.cores["muted"],
                font=self.fontes["card_titulo"],
                anchor="w",
            ).pack(anchor="w")
            tk.Label(
                bloco,
                textvariable=variavel,
                bg=self.cores["surface"],
                fg=self.cores["ink"],
                font=self.fontes["card_valor"],
                justify="left",
                wraplength=180,
                anchor="w",
            ).pack(anchor="w", pady=(6, 0))

        tk.Label(
            painel_resumo,
            text="Status da etapa",
            bg=self.cores["surface"],
            fg=self.cores["muted"],
            font=self.fontes["card_titulo"],
            anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))
        tk.Label(
            painel_resumo,
            textvariable=self.status_execucao_var,
            bg=self.cores["surface"],
            fg=self.cores["primary"],
            font=self.fontes["subtitulo"],
            justify="left",
            wraplength=320,
            anchor="w",
        ).grid(row=3, column=0, sticky="ew", pady=(6, 0))

        frame_caminhos = self._criar_superficie(self.container, background=self.cores["surface"])
        frame_caminhos.grid(row=1, column=0, sticky="ew", pady=(18, 14))
        frame_caminhos.grid_columnconfigure(1, weight=1)
        cabecalho_caminhos = self._criar_titulo_secao(
            frame_caminhos,
            "Pastas monitoradas",
            "A pasta dos setups e lida ao lado do script ou do EXE. Se precisar, ajuste aqui.",
        )
        cabecalho_caminhos.grid(row=0, column=0, columnspan=3, sticky="w")
        self._criar_selo(
            frame_caminhos,
            "Auto deteccao ativa",
            background=self.cores["tertiary_soft"],
            foreground=self.cores["tertiary"],
        ).grid(row=0, column=2, sticky="e")

        tk.Label(
            frame_caminhos,
            text="Pasta da Interface",
            bg=self.cores["surface"],
            fg=self.cores["ink"],
            font=self.fontes["card_titulo"],
        ).grid(row=2, column=0, sticky="w", pady=(18, 8))
        ttk.Entry(frame_caminhos, textvariable=self.pasta_interface_var, style="Path.TEntry").grid(
            row=2,
            column=1,
            sticky="ew",
            padx=(12, 10),
            pady=(18, 8),
        )
        ttk.Button(
            frame_caminhos,
            text="Selecionar",
            style="Soft.TButton",
            command=self._selecionar_pasta_interface,
        ).grid(row=2, column=2, sticky="ew", pady=(18, 8))

        tk.Label(
            frame_caminhos,
            text="Pasta dos setups",
            bg=self.cores["surface"],
            fg=self.cores["ink"],
            font=self.fontes["card_titulo"],
        ).grid(row=3, column=0, sticky="w", pady=(0, 2))
        ttk.Entry(frame_caminhos, textvariable=self.pasta_setups_var, style="Path.TEntry").grid(
            row=3,
            column=1,
            sticky="ew",
            padx=(12, 10),
            pady=(0, 2),
        )
        ttk.Button(
            frame_caminhos,
            text="Selecionar",
            style="Soft.TButton",
            command=self._selecionar_pasta_setups,
        ).grid(row=3, column=2, sticky="ew", pady=(0, 2))

        tk.Label(
            frame_caminhos,
            text="Se o EXE rodar dentro da pasta dist, a pasta de setups tambem e procurada na pasta pai.",
            bg=self.cores["surface"],
            fg=self.cores["muted"],
            font=self.fontes["subtitulo"],
            anchor="w",
            justify="left",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(14, 0))

        frame_status = tk.Frame(self.container, bg=self.cores["bg"])
        frame_status.grid(row=2, column=0, sticky="ew")
        cabecalho_status = self._criar_titulo_secao(
            frame_status,
            "Leitura da base atual",
            "Resumo rapido da configuracao detectada antes de qualquer alteracao.",
        )
        cabecalho_status.pack(anchor="w")

        grade_status_linha_1 = tk.Frame(frame_status, bg=self.cores["bg"])
        grade_status_linha_1.pack(fill="x", pady=(12, 10))
        grade_status_linha_2 = tk.Frame(frame_status, bg=self.cores["bg"])
        grade_status_linha_2.pack(fill="x")

        self._criar_cartao_status(grade_status_linha_1, "Banco atual", self.banco_var)
        tk.Frame(grade_status_linha_1, bg=self.cores["bg"], width=12).pack(side="left")
        self._criar_cartao_status(grade_status_linha_1, "Versao atual", self.versao_var, destaque=True)

        self._criar_cartao_status(grade_status_linha_2, "Tabela", self.tabela_var)
        tk.Frame(grade_status_linha_2, bg=self.cores["bg"], width=12).pack(side="left")
        self._criar_cartao_status(grade_status_linha_2, "Base isolada", self.isolada_var, destaque=True)

        frame_fila = self._criar_superficie(self.container)
        frame_fila.grid(row=3, column=0, sticky="nsew", pady=(14, 12))
        frame_fila.grid_columnconfigure(0, weight=1)
        frame_fila.grid_rowconfigure(2, weight=1)

        topo_fila = tk.Frame(frame_fila, bg=self.cores["surface"])
        topo_fila.grid(row=0, column=0, sticky="ew")
        topo_fila.grid_columnconfigure(0, weight=1)
        cabecalho_fila = self._criar_titulo_secao(
            topo_fila,
            "Fila de atualizacao mes a mes",
            "Sempre usa o setup mais novo de cada mes acima da versao detectada.",
        )
        cabecalho_fila.grid(row=0, column=0, sticky="w")
        tk.Label(
            topo_fila,
            textvariable=self.fila_resumo_var,
            bg=self.cores["surface_alt"],
            fg=self.cores["primary"],
            font=self.fontes["card_titulo"],
            padx=12,
            pady=7,
        ).grid(row=0, column=1, sticky="ne")

        tk.Label(
            frame_fila,
            text="Ordem prevista de execucao",
            bg=self.cores["surface"],
            fg=self.cores["muted"],
            font=self.fontes["subtitulo"],
        ).grid(row=1, column=0, sticky="w", pady=(18, 8))

        conteudo_fila = tk.Frame(
            frame_fila,
            bg=self.cores["surface_soft"],
            highlightbackground=self.cores["border"],
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        conteudo_fila.grid(row=2, column=0, sticky="nsew")
        conteudo_fila.grid_columnconfigure(0, weight=1)
        conteudo_fila.grid_rowconfigure(0, weight=1)

        self.lista_fila = tk.Listbox(
            conteudo_fila,
            height=8,
            bd=0,
            relief="flat",
            highlightthickness=0,
            selectbackground=self.cores["primary"],
            selectforeground="#ffffff",
            bg=self.cores["surface_soft"],
            fg=self.cores["ink"],
            activestyle="none",
            font=self.fontes["subtitulo"],
        )
        self.lista_fila.grid(row=0, column=0, sticky="nsew")
        barra_lista = ttk.Scrollbar(
            conteudo_fila,
            orient="vertical",
            command=self.lista_fila.yview,
            style="App.Vertical.TScrollbar",
        )
        barra_lista.grid(row=0, column=1, sticky="ns")
        self.lista_fila.configure(yscrollcommand=barra_lista.set)

        frame_execucao = self._criar_superficie(self.container)
        frame_execucao.grid(row=4, column=0, sticky="ew", pady=(0, 12))
        frame_execucao.grid_columnconfigure(0, weight=1)
        cabecalho_execucao = self._criar_titulo_secao(
            frame_execucao,
            "Execucao",
            "Se precisar, informe argumentos para os instaladores. Exemplo: /VERYSILENT /SUPPRESSMSGBOXES",
        )
        cabecalho_execucao.grid(row=0, column=0, sticky="w")

        tk.Label(
            frame_execucao,
            text="Argumentos opcionais do instalador",
            bg=self.cores["surface"],
            fg=self.cores["primary"],
            font=self.fontes["card_titulo"],
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(18, 6))

        ttk.Entry(frame_execucao, textvariable=self.argumentos_var, style="Path.TEntry").grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 12),
        )

        tk.Label(
            frame_execucao,
            text="Exemplo: /VERYSILENT /SUPPRESSMSGBOXES",
            bg=self.cores["surface"],
            fg=self.cores["muted"],
            font=self.fontes["subtitulo"],
            anchor="w",
        ).grid(row=3, column=0, sticky="w")

        frame_botoes = tk.Frame(frame_execucao, bg=self.cores["surface"])
        frame_botoes.grid(row=4, column=0, sticky="ew", pady=(16, 0))
        for indice in range(5):
            frame_botoes.grid_columnconfigure(indice, weight=1)

        self.botao_analisar = ttk.Button(
            frame_botoes,
            text="Analisar",
            style="Soft.TButton",
            command=self.analisar,
        )
        self.botao_analisar.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.botao_preparar = ttk.Button(
            frame_botoes,
            text="Preparar Base",
            style="Primary.TButton",
            command=lambda: self._iniciar_tarefa(
                "Preparando a base para atualizacao...",
                lambda: preparar_base_para_atualizacao(
                    self._pasta_interface(),
                    self._pasta_setups(),
                    logger=self._log_da_thread,
                    cancelar_evento=self.cancelamento_evento,
                ),
            ),
        )
        self.botao_preparar.grid(row=0, column=1, sticky="ew", padx=4)

        self.botao_executar = ttk.Button(
            frame_botoes,
            text="Executar Atualizacao",
            style="Accent.TButton",
            command=self._executar_atualizacao,
        )
        self.botao_executar.grid(row=0, column=2, sticky="ew", padx=4)

        self.botao_restaurar = ttk.Button(
            frame_botoes,
            text="Restaurar Nome Original",
            style="Soft.TButton",
            command=lambda: self._iniciar_tarefa(
                "Restaurando o nome original da base...",
                lambda: restaurar_nome_original_base(
                    self._pasta_interface(),
                    self._pasta_setups(),
                    logger=self._log_da_thread,
                    cancelar_evento=self.cancelamento_evento,
                ),
            ),
        )
        self.botao_restaurar.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        self.botao_parar = ttk.Button(
            frame_botoes,
            text="Parar",
            style="Soft.TButton",
            command=self._parar_execucao,
            state="disabled",
        )
        self.botao_parar.grid(row=0, column=4, sticky="ew", padx=(8, 0))

        frame_log = tk.Frame(
            self.container,
            bg=self.cores["log_bg"],
            highlightbackground=self.cores["hero_secondary"],
            highlightthickness=1,
            bd=0,
            padx=20,
            pady=18,
        )
        frame_log.grid(row=5, column=0, sticky="nsew")
        frame_log.grid_columnconfigure(0, weight=1)
        frame_log.grid_rowconfigure(1, weight=1)

        topo_log = tk.Frame(frame_log, bg=self.cores["log_bg"])
        topo_log.grid(row=0, column=0, sticky="ew")
        topo_log.grid_columnconfigure(0, weight=1)

        tk.Label(
            topo_log,
            text="Log da execucao",
            bg=self.cores["log_bg"],
            fg="#fff8ed",
            font=self.fontes["secao"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            topo_log,
            text="Tudo o que o aplicativo fizer durante a analise, preparo e atualizacao aparecera aqui.",
            bg=self.cores["log_bg"],
            fg="#bfd0d4",
            font=self.fontes["subtitulo"],
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._criar_selo(
            topo_log,
            "Tempo real",
            background=self.cores["hero_secondary"],
            foreground="#fff8ed",
        ).grid(row=0, column=1, rowspan=2, sticky="ne")

        self.caixa_log = scrolledtext.ScrolledText(
            frame_log,
            wrap="word",
            height=12,
            bg=self.cores["log_bg_soft"],
            fg=self.cores["log_fg"],
            insertbackground=self.cores["log_fg"],
            relief="flat",
            bd=0,
            padx=14,
            pady=12,
            font=self.fontes["mono"],
            selectbackground=self.cores["secondary"],
            highlightthickness=1,
            highlightbackground=self.cores["hero_secondary"],
        )
        self.caixa_log.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        self.caixa_log.configure(state="disabled")

        rodape = tk.Frame(self.container, bg=self.cores["bg"])
        rodape.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        rodape.grid_columnconfigure(0, weight=1)

        faixa_status = tk.Frame(
            rodape,
            bg=self.cores["surface_alt"],
            highlightbackground=self.cores["border"],
            highlightthickness=1,
            padx=14,
            pady=10,
        )
        faixa_status.grid(row=0, column=0, sticky="ew")
        faixa_status.grid_columnconfigure(0, weight=1)

        tk.Label(
            faixa_status,
            textvariable=self.status_execucao_var,
            bg=self.cores["surface_alt"],
            fg=self.cores["primary"],
            font=self.fontes["rodape"],
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, sticky="w")

        self.barra_progresso = ttk.Progressbar(
            faixa_status,
            style="App.Horizontal.TProgressbar",
            mode="indeterminate",
            length=240,
        )
        self.barra_progresso.grid(row=0, column=1, sticky="e")

    def _selecionar_diretorio_seguro(self, diretorio_inicial: str) -> str:
        caminho_inicial = diretorio_inicial.strip()
        pasta_inicial = Path(self._resolver_pasta_inicial(caminho_inicial))
        try:
            self.update_idletasks()
            dialogo = SeletorPastaDialog(
                self,
                "Selecione a pasta desejada.",
                pasta_inicial,
                self.cores,
                self.fontes,
            )
            return dialogo.mostrar()
        except Exception as exc:  # noqa: BLE001
            mensagem = f"Nao foi possivel abrir o seletor de pasta: {exc}"
            self._anexar_log(mensagem)
            messagebox.showerror(APP_TITULO, mensagem)
            return ""

    def _resolver_pasta_inicial(self, caminho_inicial: str) -> str:
        if not caminho_inicial:
            return str(PASTA_APLICACAO)

        texto = caminho_inicial.strip().strip('"')
        if re.fullmatch(r"[a-zA-Z]:", texto):
            texto = f"{texto}\\"

        caminho = Path(texto).expanduser()
        candidatos = [caminho, caminho.parent, PASTA_APLICACAO, Path.home()]

        for candidato in candidatos:
            try:
                if candidato.exists() and candidato.is_dir():
                    return str(candidato)
            except OSError:
                continue

        return str(PASTA_APLICACAO)

    def _selecionar_pasta_interface(self) -> None:
        selecionada = self._selecionar_diretorio_seguro(self.pasta_interface_var.get() or str(PASTA_PADRAO))
        if selecionada:
            self.pasta_interface_var.set(selecionada)

    def _selecionar_pasta_setups(self) -> None:
        selecionada = self._selecionar_diretorio_seguro(self.pasta_setups_var.get() or str(PASTA_SETUPS_PADRAO))
        if selecionada:
            self.pasta_setups_var.set(selecionada)

    def _pasta_interface(self) -> Path:
        return Path(self.pasta_interface_var.get().strip())

    def _pasta_setups(self) -> Path:
        return Path(self.pasta_setups_var.get().strip())

    def _anexar_log(self, mensagem: str) -> None:
        horario = datetime.now().strftime("%H:%M:%S")
        self.caixa_log.configure(state="normal")
        self.caixa_log.insert("end", f"[{horario}] {mensagem}\n")
        self.caixa_log.see("end")
        self.caixa_log.configure(state="disabled")

    def _log_da_thread(self, mensagem: str) -> None:
        self.fila_eventos.put(("log", mensagem))

    def _atualizar_status(self, analise: ResultadoAnalise) -> None:
        self.analise_atual = analise
        self.banco_var.set(str(analise.banco_principal.caminho))
        self.versao_var.set(formatar_versao(analise.versao_atual))
        self.tabela_var.set(analise.tabela_compilacao)
        self.isolada_var.set("SIM" if analise.base_isolada else "NAO")
        quantidade = len(analise.fila_atualizacao)
        self.fila_resumo_var.set(f"{quantidade} item(ns) na fila")
        self.status_execucao_var.set(
            f"Analise concluida. Base em {formatar_versao(analise.versao_atual)}."
        )

        self.lista_fila.delete(0, "end")
        if analise.fila_atualizacao:
            for setup in analise.fila_atualizacao:
                self.lista_fila.insert("end", setup.descricao)
        else:
            self.lista_fila.insert("end", CHAVE_FILA_VAZIA)

    def _definir_bloqueio(self, bloqueado: bool) -> None:
        self.trabalho_ativo = bloqueado
        for botao in (
            self.botao_analisar,
            self.botao_preparar,
            self.botao_executar,
            self.botao_restaurar,
        ):
            botao.configure(state="disabled" if bloqueado else "normal")
        self.botao_parar.configure(state="normal" if bloqueado else "disabled")

        if bloqueado:
            self.status_execucao_var.set("Processando. Aguarde a conclusao da etapa atual...")
            self.barra_progresso.start(10)
        else:
            self.barra_progresso.stop()

    def _parar_execucao(self) -> None:
        if not self.trabalho_ativo or self.cancelamento_evento.is_set():
            return
        self.cancelamento_evento.set()
        self.status_execucao_var.set("Parada solicitada. Encerrando a etapa atual com seguranca...")
        self._anexar_log("Parada solicitada pelo usuario.")
        self.botao_parar.configure(state="disabled")

    def _processar_fila_eventos(self) -> None:
        while True:
            try:
                tipo, valor = self.fila_eventos.get_nowait()
            except queue.Empty:
                break

            if tipo == "log":
                self._anexar_log(str(valor))
            elif tipo == "resultado":
                self._atualizar_status(valor)
            elif tipo == "erro":
                self._anexar_log(str(valor))
                self.status_execucao_var.set("Ocorreu um erro. Revise o log e os caminhos configurados.")
                messagebox.showerror(APP_TITULO, str(valor))
                self._definir_bloqueio(False)
            elif tipo == "cancelado":
                self._anexar_log(str(valor))
                self.status_execucao_var.set("Execucao interrompida pelo usuario.")
                self._definir_bloqueio(False)
            elif tipo == "fim":
                self._definir_bloqueio(False)

        self.after(150, self._processar_fila_eventos)

    def _iniciar_tarefa(self, descricao: str, funcao: Callable[[], ResultadoAnalise]) -> None:
        if self.trabalho_ativo:
            return

        self.cancelamento_evento.clear()
        self._definir_bloqueio(True)
        self.status_execucao_var.set(descricao)
        self._anexar_log(descricao)

        def worker() -> None:
            try:
                resultado = funcao()
            except OperacaoCancelada as exc:
                self.fila_eventos.put(("cancelado", str(exc)))
            except Exception as exc:  # noqa: BLE001
                self.fila_eventos.put(("erro", str(exc)))
            else:
                self.fila_eventos.put(("resultado", resultado))
            finally:
                self.fila_eventos.put(("fim", None))

        threading.Thread(target=worker, daemon=True).start()

    def analisar(self) -> None:
        self._iniciar_tarefa(
            "Analisando a base atual e os setups disponiveis...",
            lambda: analisar_atualizacao(self._pasta_interface(), self._pasta_setups()),
        )

    def _executar_atualizacao(self) -> None:
        if not messagebox.askyesno(
            APP_TITULO,
            (
                "A atualizacao vai preparar a base e executar todo o fluxo automaticamente: "
                "instalador, enters, conclusao, abertura da Interface e fechamento para seguir "
                "para a proxima versao, restauracao do nome original da base e correcao final da grid. "
                "Durante a automacao de cada setup, teclado e mouse ficarao bloqueados para evitar interferencias. "
                "Deseja continuar?"
            ),
        ):
            return

        argumentos = self.argumentos_var.get()
        self._iniciar_tarefa(
            "Executando atualizacao mes a mes...",
            lambda: executar_atualizacoes_mensais(
                self._pasta_interface(),
                self._pasta_setups(),
                argumentos_instalador=argumentos,
                preparar_antes=True,
                logger=self._log_da_thread,
                cancelar_evento=self.cancelamento_evento,
            ),
        )


def imprimir_status_cli(analise: ResultadoAnalise) -> int:
    print(formatar_resumo_analise(analise))
    return 0


def executar_cli() -> int:
    parser = argparse.ArgumentParser(description="Atualizador da base da Interface.")
    parser.add_argument(
        "acao",
        nargs="?",
        choices=("status", "preparar", "restaurar", "atualizar", "gui"),
        default="gui",
    )
    parser.add_argument("--interface", dest="pasta_interface", default=str(PASTA_PADRAO))
    parser.add_argument("--setups", dest="pasta_setups", default=str(PASTA_SETUPS_PADRAO))
    parser.add_argument("--args-installer", default="")
    parser.add_argument("--nao-preparar", action="store_true")
    argumentos = parser.parse_args()

    pasta_interface = Path(argumentos.pasta_interface)
    pasta_setups = Path(argumentos.pasta_setups)

    if acao_requer_elevacao(argumentos.acao) and not usuario_e_admin():
        if relancar_como_admin():
            return 0
        raise RuntimeError("A automacao completa precisa ser executada como Administrador.")

    if argumentos.acao == "gui":
        app = AtualizadorInterfaceApp()
        app.mainloop()
        return 0

    if argumentos.acao == "status":
        return imprimir_status_cli(analisar_atualizacao(pasta_interface, pasta_setups))

    if argumentos.acao == "preparar":
        return imprimir_status_cli(
            preparar_base_para_atualizacao(
                pasta_interface,
                pasta_setups,
                logger=print,
            )
        )

    if argumentos.acao == "restaurar":
        return imprimir_status_cli(
            restaurar_nome_original_base(
                pasta_interface,
                pasta_setups,
                logger=print,
            )
        )

    if argumentos.acao == "atualizar":
        return imprimir_status_cli(
            executar_atualizacoes_mensais(
                pasta_interface,
                pasta_setups,
                argumentos_instalador=argumentos.args_installer,
                preparar_antes=not argumentos.nao_preparar,
                logger=print,
            )
        )

    raise RuntimeError(f"Acao nao suportada: {argumentos.acao}")


def main() -> int:
    try:
        return executar_cli()
    except Exception as exc:  # noqa: BLE001
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
