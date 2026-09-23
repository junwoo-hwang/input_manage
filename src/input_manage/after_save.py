"""저장이 끝날 때마다 도는 코드. raw data 반영 코드를 main() 안에 넣으면 된다.

'기준 정보 관리' 화면에서 저장이 끝나면(S3 에 올라간 뒤) 이 파일을 따로
띄워 돌린다. 화면은 이걸 기다리지 않는다 -- 20분이 걸려도 화면은 바로
'저장 완료!' 를 띄우고 다음 일을 할 수 있다.

- print 한 것과 오류는 로그 파일에 쌓인다. 어디인지는 input_manage.py 의
  AFTER_SAVE_LOG (기본: 임시 폴더의 input_manage_after_save.log).
- 같은 파일을 연달아 저장해도 두 개가 겹쳐 돌지 않는다. 도는 중에 또 저장하면
  지금 것이 끝난 뒤 한 번 더 돈다 (그 사이 몇 번을 저장했든 한 번, 가장 최근
  저장 기준).
- AWS 키 같은 환경변수는 포털의 것을 그대로 물려받는다.
"""
import sys


def main(book: str, s3_key: str, user: str, stamp: str) -> None:
    """book   : 저장한 파일 이름, 확장자 뺀 것   (예: FAB_INPUT_ULY_r0)
    s3_key : 버킷 안의 경로                   (예: 2GAPU/input/FAB_INPUT_ULY_r0.xlsx)
             버킷은 환경변수 INPUT_S3_BUCKET (기본 G-DVC)
    user   : 저장한 사람
    stamp  : 저장한 판의 버전표 (S3 ETag)
    """
    pass


if __name__ == "__main__":
    main(*sys.argv[1:5])
