# 🌐 온라인 라이센스 검증 서버 구축 가이드

*개발 완료 후 구축 예정*

## 개요
HomeTax 시스템의 라이센스 키를 온라인으로 검증하는 서버를 Vercel을 사용하여 무료로 구축하는 방법입니다.

---

## 🚀 Vercel 기반 API 서버 구축

### **1. Vercel 장점**
- ✅ **무료 플랜**: 월 100GB 대역폭, 무제한 배포
- ✅ **서버리스**: 서버 관리 불필요
- ✅ **빠른 배포**: Git 연동으로 자동 배포
- ✅ **글로벌 CDN**: 전세계 빠른 응답속도
- ✅ **HTTPS 기본**: SSL 인증서 자동 적용

### **2. 구축 순서**

#### **Step 1: Vercel 계정 생성**
1. https://vercel.com 접속
2. GitHub 계정으로 가입
3. 새 프로젝트 생성

#### **Step 2: API 서버 코드 작성**
```python
# api/verify-license.py
import json
import hashlib
import base64
from datetime import datetime

def handler(request):
    """라이센스 키 검증 API"""
    
    # CORS 헤더 설정
    headers = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Content-Type': 'application/json'
    }
    
    if request.method == 'OPTIONS':
        return {'statusCode': 200, 'headers': headers}
    
    if request.method != 'POST':
        return {
            'statusCode': 405,
            'headers': headers,
            'body': json.dumps({'error': 'Method not allowed'})
        }
    
    try:
        # 요청 데이터 파싱
        body = json.loads(request.body)
        license_key = body.get('license_key')
        hardware_id = body.get('hardware_id')
        
        if not license_key or not hardware_id:
            return {
                'statusCode': 400,
                'headers': headers,
                'body': json.dumps({'error': 'Missing required fields'})
            }
        
        # 라이센스 검증 로직
        is_valid, message = verify_license_internal(license_key, hardware_id)
        
        # 사용 로그 기록 (선택사항)
        log_license_usage(hardware_id, license_key, is_valid)
        
        return {
            'statusCode': 200,
            'headers': headers,
            'body': json.dumps({
                'valid': is_valid,
                'message': message,
                'timestamp': datetime.now().isoformat()
            })
        }
        
    except Exception as e:
        return {
            'statusCode': 500,
            'headers': headers,
            'body': json.dumps({'error': str(e)})
        }

def verify_license_internal(license_key, hardware_id):
    """내부 라이센스 검증 로직"""
    # 여기에 라이센스 검증 로직 구현
    # (license_system.py의 verify_license_key 로직 사용)
    pass

def log_license_usage(hardware_id, license_key, is_valid):
    """라이센스 사용 로그 (선택사항)"""
    # 데이터베이스 또는 파일에 사용 기록 저장
    pass
```

#### **Step 3: 데이터베이스 연동 (선택사항)**
```python
# Vercel Postgres 또는 PlanetScale 사용
import os
import psycopg2

def get_db_connection():
    """데이터베이스 연결"""
    return psycopg2.connect(
        host=os.environ['POSTGRES_HOST'],
        database=os.environ['POSTGRES_DATABASE'],
        user=os.environ['POSTGRES_USER'],
        password=os.environ['POSTGRES_PASSWORD']
    )

def check_license_in_db(license_key, hardware_id):
    """데이터베이스에서 라이센스 확인"""
    conn = get_db_connection()
    cur = conn.cursor()
    
    # 라이센스 키가 이미 다른 PC에서 사용 중인지 확인
    cur.execute("""
        SELECT hardware_id, first_used, last_used 
        FROM license_usage 
        WHERE license_key = %s
    """, (license_key,))
    
    result = cur.fetchone()
    
    if result and result[0] != hardware_id:
        return False, "라이센스가 다른 PC에서 사용 중입니다"
    
    # 새로운 사용 기록 저장 또는 업데이트
    cur.execute("""
        INSERT INTO license_usage (license_key, hardware_id, first_used, last_used)
        VALUES (%s, %s, NOW(), NOW())
        ON CONFLICT (license_key) 
        DO UPDATE SET last_used = NOW()
    """, (license_key, hardware_id))
    
    conn.commit()
    conn.close()
    
    return True, "라이센스 검증 성공"
```

#### **Step 4: 환경 변수 설정**
```bash
# Vercel 대시보드에서 환경 변수 설정
SECRET_KEY=your_secret_key_here
POSTGRES_HOST=your_db_host
POSTGRES_DATABASE=your_db_name
POSTGRES_USER=your_db_user
POSTGRES_PASSWORD=your_db_password
```

#### **Step 5: vercel.json 설정**
```json
{
  "functions": {
    "api/verify-license.py": {
      "runtime": "python3.9"
    }
  },
  "routes": [
    {
      "src": "/api/(.*)",
      "dest": "/api/$1"
    }
  ]
}
```

### **3. 클라이언트 연동**

#### **Python 클라이언트 코드**
```python
import requests
import json

def verify_license_online(license_key, hardware_id):
    """온라인 라이센스 검증"""
    api_url = "https://your-project.vercel.app/api/verify-license"
    
    payload = {
        "license_key": license_key,
        "hardware_id": hardware_id
    }
    
    try:
        response = requests.post(
            api_url, 
            json=payload,
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            return data['valid'], data['message']
        else:
            return False, "서버 오류"
            
    except requests.exceptions.RequestException:
        # 인터넷 연결 실패시 오프라인 검증으로 대체
        return verify_license_offline(license_key, hardware_id)
```

---

## 💰 비용 분석

### **Vercel 무료 플랜**
- **API 호출**: 월 100,000회 무료
- **대역폭**: 월 100GB 무료
- **배포**: 무제한 무료
- **도메인**: 무료 서브도메인 제공

### **데이터베이스 옵션**
- **Vercel Postgres**: 무료 플랜 (소규모 프로젝트)
- **PlanetScale**: MySQL 호환, 무료 플랜
- **Supabase**: PostgreSQL, 무료 플랜
- **Firebase**: NoSQL, 무료 플랜

### **예상 월 비용**
- **소규모 (사용자 < 100명)**: **무료**
- **중규모 (사용자 < 1000명)**: **무료~$20**
- **대규모 (사용자 > 1000명)**: **$20~100**

---

## 🔒 보안 고려사항

### **1. API 보안**
- HTTPS 강제 사용
- Rate Limiting 적용
- API 키 인증 (선택사항)
- 요청 크기 제한

### **2. 라이센스 보안**
- 암호화된 라이센스 키
- 하드웨어 ID 기반 검증
- 타임스탬프 검증
- 재사용 방지

### **3. 서버 보안**
- 환경 변수로 비밀 키 관리
- 데이터베이스 접근 권한 최소화
- 로그 모니터링
- 정기적인 보안 업데이트

---

## 📊 모니터링 및 분석

### **1. 사용량 추적**
- 라이센스 검증 요청 수
- 성공/실패 비율
- 지역별 사용량
- 시간대별 패턴

### **2. 대시보드 구성**
```python
# api/dashboard.py
def get_license_stats():
    """라이센스 사용 통계"""
    return {
        "total_licenses": count_total_licenses(),
        "active_licenses": count_active_licenses(),
        "verification_requests": count_verifications_today(),
        "success_rate": calculate_success_rate()
    }
```

---

## 🚀 배포 및 운영

### **1. 자동 배포**
1. GitHub 저장소에 코드 Push
2. Vercel이 자동으로 감지하여 배포
3. 배포 완료 후 즉시 사용 가능

### **2. 모니터링**
- Vercel 대시보드에서 실시간 로그 확인
- 오류 발생시 이메일 알림
- 성능 메트릭 추적

### **3. 백업 및 복구**
- 데이터베이스 정기 백업
- 코드 버전 관리 (Git)
- 장애 복구 계획

---

## 📝 구현 체크리스트

### **개발 단계**
- [ ] Vercel 계정 생성
- [ ] API 서버 코드 작성
- [ ] 데이터베이스 설계 및 구축
- [ ] 클라이언트 연동 코드 작성
- [ ] 테스트 환경 구축

### **보안 단계**
- [ ] HTTPS 인증서 확인
- [ ] API 보안 설정
- [ ] 환경 변수 설정
- [ ] 접근 권한 설정

### **운영 단계**
- [ ] 모니터링 설정
- [ ] 백업 시스템 구축
- [ ] 문서화 완료
- [ ] 사용자 가이드 작성

---

## 📞 참고 자료

- **Vercel 공식 문서**: https://vercel.com/docs
- **Python API 가이드**: https://vercel.com/docs/functions/serverless-functions/runtimes/python
- **데이터베이스 연동**: https://vercel.com/docs/storage
- **보안 가이드**: https://vercel.com/docs/security

---

*이 가이드는 향후 온라인 라이센스 검증 서버 구축시 참고용으로 작성되었습니다.*