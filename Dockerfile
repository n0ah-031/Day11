# 취합앱 컨테이너.
#
# **워커는 하나다.** 업로드 세션과 진행 중 잡이 프로세스 메모리(dict + 임시 폴더)라,
# 워커를 늘리면 요청이 다른 워커로 가면서 "세션을 찾을 수 없습니다"가 뜬다(§6 남은 경계).
# 컨테이너를 여러 개 띄우는 것도 같은 이유로 안 된다(`docker compose up --scale app=2` 금지).
# 이 제약을 없애려면 세션·잡을 DB/Redis로 옮겨야 하고, 그건 별도 작업이다.

FROM python:3.13-slim

# 빌드 도구 없이 설치되는 wheel만 쓴다(cryptography·pillow 모두 manylinux wheel이 있다).
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# 의존성을 먼저 넣어 코드만 바뀔 때 이 레이어를 재사용한다
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 루트로 돌리지 않는다. 업로드 임시 폴더는 tempfile이 /tmp에 만들고, 그 아래만 쓴다
RUN useradd --create-home --uid 10001 app && chown -R app:app /app
USER app

# 플랫폼이 주는 PORT를 따른다(Fly·Railway·Render 모두 이 방식이다)
ENV PORT=8080
EXPOSE 8080

# --workers 를 붙이지 않는다(위 주석 참조). --proxy-headers 는 프록시 뒤에서
# 클라이언트 정보를 잃지 않기 위해서다.
CMD ["sh", "-c", "exec uvicorn server:app --host 0.0.0.0 --port ${PORT} --proxy-headers --timeout-keep-alive 75"]
