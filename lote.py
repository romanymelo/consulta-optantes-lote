from typing import Callable

from consulta import CaptchaTimeout, CnpjNaoEncontrado, ConsultaErro, PortalBloqueou
from planilha import avisos_do_portal


def _gravar(planilha, cnpj, campos, status, erro, log) -> None:
    try:
        planilha.adicionar(cnpj, campos, status, erro)
    except PermissionError as e:
        log(f"ERRO ao salvar: {e}")


def executar_lote(
    cnpjs, invalidos, consultor, planilha, log: Callable[[str], None], pasta_pdf: str | None = None
) -> None:
    for item in invalidos:
        _gravar(planilha, item, {}, "inválido", "CNPJ inválido", log)
        log(f"{item}: inválido")

    total = len(cnpjs)
    ja_avisados: set[str] = set()
    for i, cnpj in enumerate(cnpjs, 1):
        prefixo = f"[{i}/{total}] {cnpj}"
        try:
            campos = consultor.consultar(cnpj, pasta_pdf=pasta_pdf)
        except PortalBloqueou as e:
            _gravar(planilha, cnpj, {}, "bloqueado pelo portal", str(e), log)
            log(f"{prefixo}: o portal bloqueou a consulta; lote interrompido.")
            return
        except CnpjNaoEncontrado as e:
            _gravar(planilha, cnpj, {}, "não encontrado", str(e), log)
            log(f"{prefixo}: não encontrado")
        except CaptchaTimeout as e:
            _gravar(planilha, cnpj, {}, "captcha expirou", str(e), log)
            log(f"{prefixo}: captcha expirou")
        except ConsultaErro as e:
            _gravar(planilha, cnpj, {}, "erro", str(e), log)
            log(f"{prefixo}: erro — {e}")
            if "Navegador indisponível" in str(e):
                log("Navegador fechado; lote interrompido.")
                return
        else:
            _gravar(planilha, cnpj, campos, "ok", "", log)
            log(f"{prefixo}: ok")
            if pasta_pdf and "PDF" not in campos:
                log(f"{prefixo}: PDF não salvo")
            for aviso in avisos_do_portal(campos):
                if aviso not in ja_avisados:
                    ja_avisados.add(aviso)
                    log(f"{prefixo}: {aviso}")
