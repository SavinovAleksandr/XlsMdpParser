import sys, tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from mdp_converter.core import convert


def test_smoke_conversion():
    src=next((ROOT/'Исходные файлы').glob('КС Печорская ГРЭС - Ухта.xlsx'))
    out=Path(tempfile.gettempdir())/'mdp_smoke.html'
    m=convert(src,out)
    assert out.exists() and out.stat().st_size>5000
    assert len(m.schemes)>0
    html=out.read_text(encoding='utf-8')
    assert 'https://' not in html and 'http://' not in html
    assert 'Кол_во_ТГ_ПГРЭС' in html
