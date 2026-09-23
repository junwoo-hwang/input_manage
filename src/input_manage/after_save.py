"""raw data 반영 코드. main() 안에 넣는다.

포털(기준 정보 관리 화면)은 이 파일을 실행하지 않는다. 따로 1시간마다 돈다.
그래서 화면에서 저장한 값은 다음 번 돌 때 raw data 에 들어간다.
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
