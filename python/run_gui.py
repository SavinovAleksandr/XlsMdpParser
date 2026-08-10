import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parent/'src'))


def self_test() -> int:
    """Check imports needed by the GUI without opening a window."""
    import tkinter

    from mdp_converter.core import convert, convert_directory
    from mdp_converter.gui import main

    assert tkinter.TkVersion
    assert callable(convert)
    assert callable(convert_directory)
    assert callable(main)
    print("MDP Converter self-test: OK")
    return 0


if "--self-test" in sys.argv:
    raise SystemExit(self_test())

from mdp_converter.gui import main

main()
