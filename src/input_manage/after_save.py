"""raw data 반영 코드. main() 안에 넣는다.

포털(기준 정보 관리 화면)에서 저장이 끝날 때마다 포털이 이 파일을 따로
띄워 돌린다:  python after_save.py 파일이름 S3경로 저장한사람 ETag
화면은 기다리지 않는다. 끝나면 저장한 사람에게 메신저가 간다.

- 오류 없이 끝나면 '반영 완료' 메신저가 간다.
- 오류가 나서 멈추면(예외가 밖으로 나오거나 sys.exit(1)) '반영 중 오류'
  메신저가 간다. 그러니 실패를 try/except 로 삼키지 말고 그대로 두거나
  sys.exit(1) 로 끝낸다.
- print 한 것과 오류 내용은 로그 파일(환경변수 INPUT_AFTER_SAVE_LOG,
  기본 임시폴더/input_manage_after_save.log)에 쌓인다.
- 환경변수(AWS 키 등)는 포털 것을 그대로 물려받는다.
- 작업 폴더는 포털이 켜진 폴더다. 파일 경로는 절대 경로로 쓰는 게 안전하다.
"""
import sys


def main(book: str, s3_key: str, user: str, stamp: str) -> None:
    """book   : 파일 이름, 확장자 뺀 것   (예: FAB_INPUT_ULY_r0)
    s3_key : 버킷 안의 경로           (예: 2GAPU/input/FAB_INPUT_ULY_r0.xlsx)
             버킷은 환경변수 INPUT_S3_BUCKET (기본 G-DVC)
    user   : 저장한 사람
    stamp  : 저장한 판의 버전표 (S3 ETag)
    """
    pass


if __name__ == "__main__":
    main(*sys.argv[1:5])
