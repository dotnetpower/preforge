---
description: 'preforge - 엄격한 코드 생성 및 검증 규칙'
applyTo: '**'
---

# GitHub Copilot 개발 지침

## 개요

본 지침은 GitHub Copilot을 활용한 엔터프라이즈급 소프트웨어 개발을 위한 엄격한 표준을 정의합니다. 모든 코드 생성, 리뷰, 리팩토링 작업은 본 문서에 정의된 규칙을 준수해야 합니다.

## 핵심 원칙

### 1. 보안 우선 (Security First)

모든 코드는 보안을 최우선으로 고려하여 생성되어야 합니다. 의심스러운 경우 항상 더 안전한 옵션을 선택하고 그 이유를 명시해야 합니다.

**필수 준수 사항:**
- 인증 및 권한 부여 로직은 "기본 거부(deny by default)" 원칙을 따라야 함
- 모든 외부 입력은 검증 및 무해화(sanitization)되어야 함
- 민감한 데이터(API 키, 비밀번호, 연결 문자열)는 절대 하드코딩하지 않음
- 환경 변수 또는 보안 관리 서비스(예: Azure Key Vault, HashiCorp Vault)를 사용해야 함
- 기준 날짜는 2026년 1월로 지정
- 문서 내 이모지 최소화
- docs/ 폴더 내의 markdown 에 요구사항이 정의되어 있으니 참고할것.

**검증 체크리스트:**
- [ ] SQL 인젝션 방지: 파라미터화된 쿼리만 사용
- [ ] XSS 방지: 사용자 입력 출력 시 적절한 인코딩 적용
- [ ] CSRF 보호: 상태 변경 작업에 토큰 검증 구현
- [ ] 민감 데이터 암호화: 전송 중(TLS/HTTPS) 및 저장 시(AES-256) 암호화
- [ ] 최신 암호화 알고리즘 사용: bcrypt, Argon2 (MD5, SHA-1 절대 사용 금지)

### 2. 코드 품질 표준

**명명 규칙:**
- 변수, 함수, 클래스는 명확하고 의미 있는 이름을 사용해야 함
- 매직 넘버나 매직 문자열 사용 금지 (상수로 정의)
- 코드는 자체 문서화되어야 하며, 주석은 필요한 경우에만 작성

**함수 설계:**
- 단일 책임 원칙(Single Responsibility Principle) 준수
- 함수는 간결하게 유지 (이상적으로 20-30줄 이하)
- 중첩 깊이는 최대 3-4단계로 제한
- DRY(Don't Repeat Yourself) 원칙 준수

**에러 처리:**
- 적절한 수준에서 에러 처리 구현
- 의미 있는 에러 메시지 제공
- 조용한 실패(silent failure) 금지
- 입력 조기 검증(fail fast)

### 3. 테스트 필수

**테스트 커버리지:**
- 모든 중요 경로(critical path)에 대한 테스트 케이스 필수
- 경계 조건, null 값, 빈 컬렉션 등 엣지 케이스 테스트
- 테스트 이름은 테스트 내용을 명확히 설명해야 함

**테스트 구조:**
- 명확한 Arrange-Act-Assert 패턴 사용
- 테스트 간 의존성 없이 독립적으로 실행 가능해야 함
- 구체적인 assertion 사용 (generic assertTrue/assertFalse 지양)

### 4. 변경 규칙
- .md 파일의 파일명이 변경되면 sidebar 의 내용도 함께 업데이트 해줘.

### 5. 스킬 사용 규칙
TBD

---

## Python 개발 지침

### 코드 스타일

**PEP 8 준수:**
```python
# 필수 사항
# - 4 스페이스 들여쓰기
# - 최대 줄 길이 79자
# - 함수와 클래스 사이 적절한 빈 줄
```

**타입 힌트 및 문서화:**
```python
from typing import List, Dict, Optional

def calculate_total_price(
    items: List[Dict[str, float]], 
    tax_rate: float = 0.1
) -> float:
    """
    주어진 항목들의 총 가격을 세금을 포함하여 계산합니다.
    
    Args:
        items: 가격 정보가 포함된 항목 딕셔너리 리스트
        tax_rate: 세율 (기본값: 0.1)
    
    Returns:
        세금이 포함된 총 가격
    
    Raises:
        ValueError: items가 비어있거나 tax_rate가 음수인 경우
    
    Example:
        >>> items = [{"price": 100}, {"price": 200}]
        >>> calculate_total_price(items, 0.1)
        330.0
    """
    if not items:
        raise ValueError("항목 리스트가 비어있습니다")
    if tax_rate < 0:
        raise ValueError("세율은 0 이상이어야 합니다")
    
    subtotal = sum(item["price"] for item in items)
    return subtotal * (1 + tax_rate)
```

### 보안 규칙

**SQL 인젝션 방지:**
```python
# ❌ 절대 금지
def get_user_bad(email: str):
    query = f"SELECT * FROM users WHERE email = '{email}'"
    return db.execute(query)

# ✅ 올바른 방법
def get_user_good(email: str):
    query = "SELECT * FROM users WHERE email = ?"
    return db.execute(query, (email,))
```

**비밀 정보 관리:**
```python
# ❌ 절대 금지
API_KEY = "sk_live_abc123xyz789"

# ✅ 올바른 방법
import os
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

# 환경 변수 사용
API_KEY = os.getenv("API_KEY")
if not API_KEY:
    raise ValueError("API_KEY 환경 변수가 설정되지 않았습니다")

# 또는 Azure Key Vault 사용
credential = DefaultAzureCredential()
client = SecretClient(vault_url="https://myvault.vault.azure.net/", credential=credential)
API_KEY = client.get_secret("api-key").value
```

### 에러 처리

```python
# ❌ 잘못된 예
def process_data(data):
    try:
        result = complex_operation(data)
    except:
        pass  # 조용한 실패 - 절대 금지

# ✅ 올바른 예
from typing import Optional
import logging

logger = logging.getLogger(__name__)

def process_data(data: dict) -> Optional[dict]:
    """
    데이터를 처리하고 결과를 반환합니다.
    
    Args:
        data: 처리할 데이터 딕셔너리
    
    Returns:
        처리된 결과 또는 None
    
    Raises:
        ValueError: 데이터가 유효하지 않은 경우
        ProcessingError: 처리 중 오류가 발생한 경우
    """
    if not data or not isinstance(data, dict):
        raise ValueError(f"유효하지 않은 데이터 형식: {type(data)}")
    
    try:
        result = complex_operation(data)
        return result
    except KeyError as e:
        logger.error(f"필수 키 누락: {e}")
        raise ValueError(f"필수 데이터 누락: {e}")
    except Exception as e:
        logger.exception(f"데이터 처리 실패: {e}")
        raise ProcessingError(f"처리 중 오류 발생: {e}") from e
```

### 테스트

```python
import pytest
from decimal import Decimal

def test_calculate_total_price_with_valid_items():
    """유효한 항목으로 총 가격이 정확히 계산되는지 검증"""
    # Arrange
    items = [
        {"price": 100.0},
        {"price": 200.0}
    ]
    tax_rate = 0.1
    
    # Act
    result = calculate_total_price(items, tax_rate)
    
    # Assert
    assert result == 330.0

def test_calculate_total_price_raises_error_for_empty_items():
    """빈 항목 리스트로 호출 시 ValueError 발생 검증"""
    # Arrange
    items = []
    tax_rate = 0.1
    
    # Act & Assert
    with pytest.raises(ValueError, match="항목 리스트가 비어있습니다"):
        calculate_total_price(items, tax_rate)

def test_calculate_total_price_raises_error_for_negative_tax():
    """음수 세율로 호출 시 ValueError 발생 검증"""
    # Arrange
    items = [{"price": 100.0}]
    tax_rate = -0.1
    
    # Act & Assert
    with pytest.raises(ValueError, match="세율은 0 이상이어야 합니다"):
        calculate_total_price(items, tax_rate)
```

---

## .NET/C# 개발 지침

### 코드 스타일

**최신 C# 기능 사용 (C# 12+):**
```csharp
// 필수: 최신 C# 기능 활용
// - File-scoped namespace
// - Primary constructors
// - Record types
// - Pattern matching
// - Nullable reference types
```

**명명 규칙:**
```csharp
// PascalCase: 클래스, 메서드, 프로퍼티, 인터페이스
public class OrderService { }
public interface IOrderRepository { }
public void ProcessOrder() { }
public int TotalAmount { get; set; }

// camelCase: private 필드, 로컬 변수
private readonly ILogger _logger;
private int itemCount;
```

**Nullable Reference Types:**
```csharp
#nullable enable

public class UserService
{
    private readonly IUserRepository _repository;
    
    public UserService(IUserRepository repository)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }
    
    // 명시적 null 가능성 선언
    public User? FindUser(string userId)
    {
        if (string.IsNullOrWhiteSpace(userId))
        {
            throw new ArgumentException("사용자 ID는 필수입니다", nameof(userId));
        }
        
        return _repository.FindById(userId);
    }
    
    // null 체크는 is null/is not null 사용
    public void ProcessUser(User? user)
    {
        if (user is null)
        {
            throw new ArgumentNullException(nameof(user));
        }
        
        // user는 여기서 null이 아님
        Console.WriteLine(user.Name);
    }
}
```

### 보안 규칙

**SQL 인젝션 방지 (Entity Framework Core):**
```csharp
// ❌ 절대 금지 - 문자열 연결
public async Task<User> GetUserBadAsync(string email)
{
    var query = $"SELECT * FROM Users WHERE Email = '{email}'";
    return await _context.Users.FromSqlRaw(query).FirstOrDefaultAsync();
}

// ✅ 올바른 방법 - 파라미터화된 쿼리
public async Task<User?> GetUserGoodAsync(string email)
{
    return await _context.Users
        .Where(u => u.Email == email)
        .FirstOrDefaultAsync();
}

// ✅ 또는 FromSqlInterpolated 사용
public async Task<User?> GetUserSafeAsync(string email)
{
    return await _context.Users
        .FromSqlInterpolated($"SELECT * FROM Users WHERE Email = {email}")
        .FirstOrDefaultAsync();
}
```

**비밀 정보 관리:**
```csharp
// ❌ 절대 금지
public class ApiSettings
{
    public const string ApiKey = "sk_live_abc123xyz789";
}

// ✅ 올바른 방법 - Configuration 사용
public class ApiSettings
{
    public string ApiKey { get; set; } = string.Empty;
}

// Program.cs 또는 Startup.cs
public void ConfigureServices(IServiceCollection services)
{
    // appsettings.json 또는 환경 변수에서 로드
    services.Configure<ApiSettings>(Configuration.GetSection("Api"));
    
    // 또는 Azure Key Vault 사용
    var keyVaultUrl = Configuration["KeyVault:Url"];
    var credential = new DefaultAzureCredential();
    var client = new SecretClient(new Uri(keyVaultUrl), credential);
    var apiKey = await client.GetSecretAsync("api-key");
}
```

### 에러 처리 및 검증

```csharp
public class OrderService
{
    private readonly IOrderRepository _repository;
    private readonly ILogger<OrderService> _logger;
    
    public OrderService(
        IOrderRepository repository, 
        ILogger<OrderService> logger)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }
    
    /// <summary>
    /// 주문을 처리합니다.
    /// </summary>
    /// <param name="orderId">주문 ID</param>
    /// <param name="cancellationToken">취소 토큰</param>
    /// <returns>처리된 주문</returns>
    /// <exception cref="ArgumentException">orderId가 유효하지 않은 경우</exception>
    /// <exception cref="OrderNotFoundException">주문을 찾을 수 없는 경우</exception>
    /// <exception cref="OrderProcessingException">주문 처리 중 오류가 발생한 경우</exception>
    public async Task<Order> ProcessOrderAsync(
        string orderId, 
        CancellationToken cancellationToken = default)
    {
        // 입력 검증
        if (string.IsNullOrWhiteSpace(orderId))
        {
            throw new ArgumentException("주문 ID는 필수입니다", nameof(orderId));
        }
        
        try
        {
            var order = await _repository.GetByIdAsync(orderId, cancellationToken);
            
            if (order is null)
            {
                _logger.LogWarning("주문을 찾을 수 없음: {OrderId}", orderId);
                throw new OrderNotFoundException($"주문 ID {orderId}를 찾을 수 없습니다");
            }
            
            // 비즈니스 로직
            order.Process();
            await _repository.UpdateAsync(order, cancellationToken);
            
            _logger.LogInformation("주문 처리 완료: {OrderId}", orderId);
            return order;
        }
        catch (OrderNotFoundException)
        {
            throw; // 이미 로깅되었으므로 다시 던짐
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "주문 처리 실패: {OrderId}", orderId);
            throw new OrderProcessingException($"주문 처리 중 오류 발생: {ex.Message}", ex);
        }
    }
}
```

### 테스트

```csharp
public class OrderServiceTests
{
    private readonly Mock<IOrderRepository> _repositoryMock;
    private readonly Mock<ILogger<OrderService>> _loggerMock;
    private readonly OrderService _service;
    
    public OrderServiceTests()
    {
        _repositoryMock = new Mock<IOrderRepository>();
        _loggerMock = new Mock<ILogger<OrderService>>();
        _service = new OrderService(_repositoryMock.Object, _loggerMock.Object);
    }
    
    [Fact]
    public async Task ProcessOrderAsync_ValidOrderId_ReturnsProcessedOrder()
    {
        // Arrange
        var orderId = "ORDER-123";
        var order = new Order { Id = orderId, Status = OrderStatus.Pending };
        _repositoryMock.Setup(r => r.GetByIdAsync(orderId, default))
            .ReturnsAsync(order);
        
        // Act
        var result = await _service.ProcessOrderAsync(orderId);
        
        // Assert
        result.Should().NotBeNull();
        result.Status.Should().Be(OrderStatus.Processed);
        _repositoryMock.Verify(r => r.UpdateAsync(order, default), Times.Once);
    }
    
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task ProcessOrderAsync_InvalidOrderId_ThrowsArgumentException(string orderId)
    {
        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(
            () => _service.ProcessOrderAsync(orderId));
    }
    
    [Fact]
    public async Task ProcessOrderAsync_OrderNotFound_ThrowsOrderNotFoundException()
    {
        // Arrange
        var orderId = "NONEXISTENT";
        _repositoryMock.Setup(r => r.GetByIdAsync(orderId, default))
            .ReturnsAsync((Order?)null);
        
        // Act & Assert
        var exception = await Assert.ThrowsAsync<OrderNotFoundException>(
            () => _service.ProcessOrderAsync(orderId));
        
        exception.Message.Should().Contain(orderId);
    }
}
```

---

## 컨테이너화 및 인프라 지침

### Docker 모범 사례

**멀티스테이지 빌드 (필수):**
```dockerfile
# Python 예제
# Stage 1: 빌드 환경
FROM python:3.12-slim AS builder

WORKDIR /app

# 의존성 파일만 먼저 복사 (캐싱 최적화)
COPY requirements.txt .
RUN pip install --user --no-cache-dir -r requirements.txt

# 소스 코드 복사
COPY src/ ./src/

# Stage 2: 프로덕션 환경
FROM python:3.12-slim

WORKDIR /app

# 빌드 단계에서 설치한 패키지만 복사
COPY --from=builder /root/.local /root/.local
COPY --from=builder /app/src ./src

# 비루트 사용자 생성 및 전환
RUN addgroup --system appgroup && \
    adduser --system --ingroup appgroup appuser && \
    chown -R appuser:appgroup /app

USER appuser

# 환경 변수로 설정 관리
ENV PATH=/root/.local/bin:$PATH \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# 포트 노출 (문서화 목적)
EXPOSE 8000

# 헬스체크
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')" || exit 1

CMD ["python", "-m", "src.main"]
```

```dockerfile
# .NET 예제
# Stage 1: SDK 이미지로 빌드
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build

WORKDIR /src
COPY ["MyApp.csproj", "./"]
RUN dotnet restore "MyApp.csproj"

COPY . .
RUN dotnet build "MyApp.csproj" -c Release -o /app/build

# Stage 2: 게시
FROM build AS publish
RUN dotnet publish "MyApp.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Stage 3: 런타임 이미지
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final

WORKDIR /app

# 비루트 사용자 생성
RUN groupadd -r appgroup && useradd -r -g appgroup appuser && \
    chown -R appuser:appgroup /app

USER appuser

COPY --from=publish /app/publish .

EXPOSE 8080

# 헬스체크
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD curl --fail http://localhost:8080/health || exit 1

ENTRYPOINT ["dotnet", "MyApp.dll"]
```

**.dockerignore (필수):**
```dockerignore
# 버전 관리
.git
.gitignore
.gitattributes

# 의존성 (컨테이너 내부에서 설치)
**/node_modules
**/venv
**/__pycache__
**/bin
**/obj

# 빌드 아티팩트
**/dist
**/build
**/*.o
**/*.so

# 개발 파일
.env*
*.log
coverage/
.pytest_cache/

# IDE 파일
.vscode/
.idea/
*.swp

# OS 파일
.DS_Store
Thumbs.db

# 문서
*.md
docs/

# 테스트
**/tests/
**/*test.py
**/*Test.cs
```

### Kubernetes 배포 규칙

**배포 매니페스트 (검증 필수):**
```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: myapp
  namespace: production
  labels:
    app: myapp
    version: v1.0.0
spec:
  replicas: 3
  selector:
    matchLabels:
      app: myapp
  strategy:
    type: RollingUpdate
    rollingUpdate:
      maxSurge: 1
      maxUnavailable: 0
  template:
    metadata:
      labels:
        app: myapp
        version: v1.0.0
    spec:
      # 보안 컨텍스트
      securityContext:
        runAsNonRoot: true
        runAsUser: 1000
        fsGroup: 2000
      
      containers:
      - name: myapp
        image: myregistry.azurecr.io/myapp:v1.0.0
        imagePullPolicy: Always
        
        ports:
        - containerPort: 8080
          name: http
          protocol: TCP
        
        # 필수: 리소스 제한
        resources:
          requests:
            cpu: "100m"
            memory: "128Mi"
          limits:
            cpu: "500m"
            memory: "512Mi"
        
        # 필수: 헬스체크
        livenessProbe:
          httpGet:
            path: /health/live
            port: 8080
          initialDelaySeconds: 30
          periodSeconds: 10
          timeoutSeconds: 5
          failureThreshold: 3
        
        readinessProbe:
          httpGet:
            path: /health/ready
            port: 8080
          initialDelaySeconds: 10
          periodSeconds: 5
          timeoutSeconds: 3
          failureThreshold: 3
        
        # 환경 변수
        env:
        - name: ASPNETCORE_ENVIRONMENT
          value: "Production"
        - name: DATABASE_CONNECTION
          valueFrom:
            secretKeyRef:
              name: myapp-secrets
              key: database-connection
        
        # 보안 컨텍스트
        securityContext:
          allowPrivilegeEscalation: false
          readOnlyRootFilesystem: true
          capabilities:
            drop:
            - ALL
        
        # 볼륨 마운트
        volumeMounts:
        - name: tmp
          mountPath: /tmp
        - name: config
          mountPath: /app/config
          readOnly: true
      
      volumes:
      - name: tmp
        emptyDir: {}
      - name: config
        configMap:
          name: myapp-config

---
apiVersion: v1
kind: Service
metadata:
  name: myapp-service
  namespace: production
spec:
  selector:
    app: myapp
  ports:
  - protocol: TCP
    port: 80
    targetPort: 8080
  type: ClusterIP

---
apiVersion: networking.k8s.io/v1
kind: NetworkPolicy
metadata:
  name: myapp-netpol
  namespace: production
spec:
  podSelector:
    matchLabels:
      app: myapp
  policyTypes:
  - Ingress
  - Egress
  ingress:
  - from:
    - namespaceSelector:
        matchLabels:
          name: ingress-nginx
    ports:
    - protocol: TCP
      port: 8080
  egress:
  - to:
    - namespaceSelector:
        matchLabels:
          name: kube-system
    ports:
    - protocol: TCP
      port: 53
  - to:
    - podSelector:
        matchLabels:
          app: database
    ports:
    - protocol: TCP
      port: 5432
```

---

## DevOps 및 CI/CD 원칙

### CALMS 프레임워크

**Culture (문화):**
- 협업과 공유 책임 문화 조성
- 무책임(blameless) 사후 분석 수행
- 지속적 학습과 개선 추구

**Automation (자동화):**
- CI/CD 파이프라인 자동화 필수
- Infrastructure as Code (IaC) 사용
- 보안 스캔 자동화 (SAST, DAST, SCA)
- 자동화된 테스트 실행

**Lean (린):**
- 작은 배치 크기 유지 (작은 PR, 빈번한 배포)
- 가치 흐름 최적화
- 낭비 제거

**Measurement (측정):**
- DORA 메트릭 추적 필수
  - 배포 빈도 (Deployment Frequency)
  - 변경 리드 타임 (Lead Time for Changes)
  - 변경 실패율 (Change Failure Rate)
  - 평균 복구 시간 (MTTR)

**Sharing (공유):**
- 지식과 도구 공유
- 명확한 문서화
- 크로스 펑셔널 팀 구성

### GitHub Actions 워크플로우 예제

```yaml
name: CI/CD Pipeline

on:
  push:
    branches: [ main, develop ]
  pull_request:
    branches: [ main ]

env:
  DOTNET_VERSION: '8.0'
  PYTHON_VERSION: '3.12'

jobs:
  security-scan:
    name: Security Scanning
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Run Trivy vulnerability scanner
        uses: aquasecurity/trivy-action@master
        with:
          scan-type: 'fs'
          scan-ref: '.'
          format: 'sarif'
          output: 'trivy-results.sarif'
      
      - name: Upload Trivy results to GitHub Security
        uses: github/codeql-action/upload-sarif@v2
        with:
          sarif_file: 'trivy-results.sarif'

  test-python:
    name: Python Tests
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: ${{ env.PYTHON_VERSION }}
          cache: 'pip'
      
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          pip install pytest pytest-cov
      
      - name: Run tests with coverage
        run: |
          pytest --cov=src --cov-report=xml --cov-report=term
      
      - name: Upload coverage to Codecov
        uses: codecov/codecov-action@v3
        with:
          file: ./coverage.xml
          fail_ci_if_error: true

  test-dotnet:
    name: .NET Tests
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Setup .NET
        uses: actions/setup-dotnet@v3
        with:
          dotnet-version: ${{ env.DOTNET_VERSION }}
      
      - name: Restore dependencies
        run: dotnet restore
      
      - name: Build
        run: dotnet build --no-restore --configuration Release
      
      - name: Test
        run: |
          dotnet test --no-build --configuration Release \
            --logger "trx" \
            --collect:"XPlat Code Coverage" \
            --results-directory ./TestResults
      
      - name: Upload test results
        uses: actions/upload-artifact@v3
        if: always()
        with:
          name: test-results
          path: ./TestResults

  build-and-push:
    name: Build and Push Docker Image
    needs: [security-scan, test-python, test-dotnet]
    runs-on: ubuntu-latest
    if: github.ref == 'refs/heads/main'
    steps:
      - uses: actions/checkout@v4
      
      - name: Log in to Azure Container Registry
        uses: docker/login-action@v3
        with:
          registry: ${{ secrets.ACR_REGISTRY }}
          username: ${{ secrets.ACR_USERNAME }}
          password: ${{ secrets.ACR_PASSWORD }}
      
      - name: Extract metadata
        id: meta
        uses: docker/metadata-action@v5
        with:
          images: ${{ secrets.ACR_REGISTRY }}/myapp
          tags: |
            type=ref,event=branch
            type=sha,prefix={{branch}}-
            type=semver,pattern={{version}}
      
      - name: Build and push
        uses: docker/build-push-action@v5
        with:
          context: .
          push: true
          tags: ${{ steps.meta.outputs.tags }}
          labels: ${{ steps.meta.outputs.labels }}
          cache-from: type=gha
          cache-to: type=gha,mode=max
      
      - name: Scan image
        uses: aquasecurity/trivy-action@master
        with:
          image-ref: ${{ secrets.ACR_REGISTRY }}/myapp:${{ steps.meta.outputs.version }}
          format: 'sarif'
          output: 'trivy-image-results.sarif'
      
      - name: Upload image scan results
        uses: github/codeql-action/upload-sarif@v2
        with:
          sarif_file: 'trivy-image-results.sarif'
```

---

## 코드 리뷰 지침

### 우선순위

**🔴 CRITICAL (병합 차단):**
- 보안 취약점 (SQL 인젝션, XSS, CSRF 등)
- 데이터 손실 위험
- 인증/권한 부여 문제
- 비밀 정보 노출

**🟡 IMPORTANT (논의 필요):**
- SOLID 원칙 위반
- 중요 경로의 테스트 누락
- 성능 병목 (N+1 쿼리, 메모리 누수)
- 아키텍처 패턴 이탈

**🟢 SUGGESTION (개선 제안):**
- 가독성 향상
- 네이밍 개선
- 코드 중복 제거
- 문서화 보완

### 리뷰 코멘트 형식

```markdown
**[우선순위] 카테고리: 간략한 제목**

상세 설명

**문제점:**
현재 코드의 문제점 설명

**제안 사항:**
```언어
// 개선된 코드 예시
```

**참고 자료:** [관련 문서 링크]
```

### 리뷰 체크리스트

**보안:**
- [ ] 민감 데이터가 코드나 로그에 노출되지 않음
- [ ] 모든 사용자 입력이 검증됨
- [ ] SQL 인젝션 취약점 없음
- [ ] 인증 및 권한 부여가 적절히 구현됨
- [ ] 의존성이 최신이며 알려진 취약점 없음

**코드 품질:**
- [ ] 일관된 코드 스타일 준수
- [ ] 의미 있는 변수/함수 이름 사용
- [ ] 함수가 간결하고 단일 책임 준수
- [ ] 코드 중복 없음
- [ ] 적절한 에러 처리

**테스트:**
- [ ] 새 코드에 적절한 테스트 커버리지
- [ ] 엣지 케이스 및 에러 시나리오 테스트
- [ ] 테스트가 독립적이고 결정적임

**성능:**
- [ ] 명백한 성능 문제 없음 (N+1, 메모리 누수)
- [ ] 적절한 캐싱 사용
- [ ] 효율적인 알고리즘 및 자료구조

**문서화:**
- [ ] 공개 API 문서화
- [ ] 복잡한 로직에 설명 주석
- [ ] README 업데이트 (필요시)

---

## 검증 프로세스

### 코드 생성 시 필수 검증

1. **보안 검증**
   - 모든 외부 입력 검증 확인
   - 비밀 정보 하드코딩 여부 확인
   - 인증/권한 부여 로직 검증

2. **품질 검증**
   - 코드 스타일 가이드 준수 확인
   - 명명 규칙 준수 확인
   - 함수 복잡도 확인 (최대 20-30줄)

3. **테스트 검증**
   - 단위 테스트 작성 확인
   - 엣지 케이스 테스트 확인
   - 테스트 커버리지 확인

4. **문서화 검증**
   - 함수/메서드 문서화 확인
   - 복잡한 로직 주석 확인
   - API 문서 업데이트 확인

### 배포 전 체크리스트

- [ ] 모든 테스트 통과
- [ ] 보안 스캔 통과 (SAST, DAST, 의존성 스캔)
- [ ] 코드 리뷰 승인
- [ ] 문서 업데이트
- [ ] 환경 변수 및 시크릿 설정 확인
- [ ] 모니터링 및 로깅 설정 확인
- [ ] 롤백 계획 수립
- [ ] 성능 테스트 완료 (필요시)

---

## 로깅 및 모니터링

### 구조화된 로깅

**Python 예제:**
```python
import logging
import json
from datetime import datetime

class StructuredLogger:
    def __init__(self, name: str):
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.INFO)
        
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(handler)
    
    def log(self, level: str, message: str, **kwargs):
        log_entry = {
            "timestamp": datetime.utcnow().isoformat(),
            "level": level,
            "message": message,
            "logger": self.logger.name,
            **kwargs
        }
        self.logger.log(
            getattr(logging, level.upper()),
            json.dumps(log_entry)
        )

# 사용 예
logger = StructuredLogger(__name__)
logger.log("info", "사용자 로그인 성공", user_id="12345", ip_address="192.168.1.1")
logger.log("error", "데이터베이스 연결 실패", error="Connection timeout", retry_count=3)
```

**C# 예제:**
```csharp
public class OrderService
{
    private readonly ILogger<OrderService> _logger;
    
    public async Task<Order> ProcessOrderAsync(string orderId)
    {
        using var scope = _logger.BeginScope(new Dictionary<string, object>
        {
            ["OrderId"] = orderId,
            ["CorrelationId"] = Activity.Current?.Id ?? Guid.NewGuid().ToString()
        });
        
        _logger.LogInformation("주문 처리 시작");
        
        try
        {
            var order = await _repository.GetByIdAsync(orderId);
            
            _logger.LogInformation(
                "주문 조회 완료: {OrderStatus}, {ItemCount}개 항목", 
                order.Status, 
                order.Items.Count);
            
            order.Process();
            await _repository.UpdateAsync(order);
            
            _logger.LogInformation("주문 처리 완료");
            return order;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex, 
                "주문 처리 실패: {ErrorMessage}", 
                ex.Message);
            throw;
        }
    }
}
```

### 헬스체크 엔드포인트 (필수)

**Python (FastAPI):**
```python
from fastapi import FastAPI, status
from fastapi.responses import JSONResponse
import asyncpg

app = FastAPI()

@app.get("/health/live")
async def liveness():
    """컨테이너가 살아있는지 확인"""
    return {"status": "alive"}

@app.get("/health/ready")
async def readiness(db: asyncpg.Pool):
    """애플리케이션이 요청을 처리할 준비가 되었는지 확인"""
    try:
        # 데이터베이스 연결 확인
        async with db.acquire() as conn:
            await conn.fetchval("SELECT 1")
        
        return {
            "status": "ready",
            "checks": {
                "database": "ok"
            }
        }
    except Exception as e:
        return JSONResponse(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            content={
                "status": "not_ready",
                "checks": {
                    "database": f"error: {str(e)}"
                }
            }
        )
```


---

## 성능 최적화

### 데이터베이스 최적화

**N+1 쿼리 방지 (Python):**
```python
# ❌ 잘못된 예 - N+1 쿼리
def get_users_with_orders_bad():
    users = session.query(User).all()
    for user in users:
        # 각 사용자마다 별도의 쿼리 실행
        orders = session.query(Order).filter_by(user_id=user.id).all()

# ✅ 올바른 예 - JOIN 사용
from sqlalchemy.orm import joinedload

def get_users_with_orders_good():
    users = session.query(User)\
        .options(joinedload(User.orders))\
        .all()
```

**N+1 쿼리 방지 (C#):**
```csharp
// ❌ 잘못된 예 - N+1 쿼리
public async Task<List<User>> GetUsersWithOrdersBadAsync()
{
    var users = await _context.Users.ToListAsync();
    foreach (var user in users)
    {
        // 각 사용자마다 별도의 쿼리 실행
        user.Orders = await _context.Orders
            .Where(o => o.UserId == user.Id)
            .ToListAsync();
    }
    return users;
}

// ✅ 올바른 예 - Include 사용
public async Task<List<User>> GetUsersWithOrdersGoodAsync()
{
    return await _context.Users
        .Include(u => u.Orders)
        .ThenInclude(o => o.Items)
        .ToListAsync();
}
```

### 캐싱 전략

**Python (Redis):**
```python
import redis
import json
from functools import wraps
from typing import Callable, Any

redis_client = redis.Redis(host='localhost', port=6379, decode_responses=True)

def cache_result(ttl: int = 300):
    """결과를 Redis에 캐싱하는 데코레이터"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        async def wrapper(*args, **kwargs) -> Any:
            # 캐시 키 생성
            cache_key = f"{func.__name__}:{args}:{kwargs}"
            
            # 캐시 확인
            cached = redis_client.get(cache_key)
            if cached:
                return json.loads(cached)
            
            # 함수 실행
            result = await func(*args, **kwargs)
            
            # 결과 캐싱
            redis_client.setex(
                cache_key,
                ttl,
                json.dumps(result)
            )
            
            return result
        return wrapper
    return decorator

@cache_result(ttl=600)
async def get_product_details(product_id: str) -> dict:
    # 데이터베이스에서 제품 정보 조회
    return await fetch_from_database(product_id)
```

**C# (IMemoryCache):**
```csharp
public class ProductService
{
    private readonly IMemoryCache _cache;
    private readonly IProductRepository _repository;
    private readonly ILogger<ProductService> _logger;
    
    public async Task<Product> GetProductAsync(string productId)
    {
        var cacheKey = $"product:{productId}";
        
        if (_cache.TryGetValue(cacheKey, out Product cachedProduct))
        {
            _logger.LogDebug("캐시에서 제품 반환: {ProductId}", productId);
            return cachedProduct;
        }
        
        var product = await _repository.GetByIdAsync(productId);
        
        var cacheOptions = new MemoryCacheEntryOptions()
            .SetSlidingExpiration(TimeSpan.FromMinutes(10))
            .SetAbsoluteExpiration(TimeSpan.FromHours(1))
            .RegisterPostEvictionCallback((key, value, reason, state) =>
            {
                _logger.LogDebug(
                    "캐시 제거: {Key}, 이유: {Reason}", 
                    key, 
                    reason);
            });
        
        _cache.Set(cacheKey, product, cacheOptions);
        
        return product;
    }
}
```

---

## 결론

본 지침은 GitHub Copilot을 활용한 엔터프라이즈급 소프트웨어 개발을 위한 기준을 제시합니다. 모든 코드는 다음 원칙을 준수해야 합니다:

1. **보안 우선**: 모든 결정에서 보안을 최우선으로 고려
2. **품질 보장**: 높은 코드 품질 및 테스트 커버리지 유지
3. **자동화**: 반복 작업의 자동화로 효율성 향상
4. **측정 및 개선**: 메트릭 기반 지속적 개선
5. **협업 및 공유**: 지식과 베스트 프랙티스 공유

이러한 원칙을 준수함으로써 안전하고, 확장 가능하며, 유지보수가 용이한 소프트웨어를 개발할 수 있습니다.

---

**문서 버전:** 1.0.0  
**최종 업데이트:** 2025년 12월  
**적용 대상:** Python 3.12+, .NET 8.0+, Docker, Kubernetes
