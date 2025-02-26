import logging

# 로거 설정
logger = logging.getLogger('MyLogger')
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()  # 콘솔로 출력
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)

# 다른 모듈에서 로깅
def some_function():
    logger.info("작업 시작")
    logger.debug("디버그 정보")
    logger.warning("경고 메시지")
    logger.error("에러 메시지")