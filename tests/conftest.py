# 저장소 루트에 있는 app 모듈을 테스트에서 불러올 수 있게 합니다.
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
