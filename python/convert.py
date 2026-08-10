import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parent/'src'))
from mdp_converter.cli import main
raise SystemExit(main())
