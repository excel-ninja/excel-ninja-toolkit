# ExcelNinja 이슈 수정 계획

> 작성일: 2025-01-11  
> 프로젝트: ExcelNinja (io.github.excel-ninja:excelNinja)  
> 현재 버전: 1.0.1

---

## 📋 목차

1. [이슈 요약](#이슈-요약)
2. [현재 상태](#현재-상태)
3. [Phase 1: Critical 이슈 수정](#phase-1-critical-이슈-수정)
4. [Phase 2: High 이슈 수정](#phase-2-high-이슈-수정)
5. [Phase 3: Medium 이슈 수정](#phase-3-medium-이슈-수정)
6. [테스트 계획](#테스트-계획)
7. [릴리즈 체크리스트](#릴리즈-체크리스트)

---

## 이슈 요약

| 우선순위 | 이슈 | 영향도 | 예상 소요 |
|---------|------|--------|----------|
| 🚨 Critical | Java 버전 호환성 | 런타임 오류 | 30분 |
| 🚨 Critical | ChunkIterator 스레드 누수 | 메모리/리소스 누수 | 1시간 |
| 🚨 Critical | EntityMetadata 캐시 누수 | 메모리 누수 | 1시간 |
| ⚠️ High | NinjaExcel Thread Safety | 동시성 버그 | 2시간 |
| ⚠️ High | 헤더 셀 타입 예외 처리 | 런타임 오류 | 30분 |
| ⚠️ High | OPCPackage InputStream 소비 | 예상치 못한 동작 | 1시간 |
| 📝 Medium | README 버전 불일치 | 문서 오류 | 10분 |
| 📝 Medium | 로깅 레벨 일관성 | 가독성 | 30분 |
| 📝 Medium | Streaming 임계값 불일치 | 예상치 못한 동작 | 20분 |

**총 예상 소요 시간: 약 7시간**

## 현재 상태

> 검토일: 2026-03-12

- [x] 1.1 Java 버전 호환성: `build.gradle` toolchain을 Java 17 LTS로 조정하고 Java 8 bytecode(`options.release = 8`) 유지
- [x] 1.2 ChunkIterator 스레드 리소스 누수: `close()`에 interrupt, join, queue clear, stream close 반영
- [x] 1.3 EntityMetadata 캐시 누수: LRU eviction(max 1000) 적용, 캐시 크기 상한 및 동시 lookup 테스트 추가
- [x] 2.1 NinjaExcel Thread Safety: reader/writer stateless 공유 유지, 동시 read/write 회귀 테스트 추가
- [x] 2.2 헤더 셀 타입 예외 처리: 숫자/불리언/수식 헤더 변환 지원 및 회귀 테스트 추가
- [x] 2.3 OPCPackage InputStream 소비 문제: 문서화 및 명시적 close 처리 반영
- [x] 3.1 README 버전 불일치: README 의존성 버전을 현재 프로젝트 버전과 동기화
- [x] 3.2 로깅 레벨 일관성: 주요 `SEVERE` 로그를 예외 stacktrace 포함 방식으로 통일하고 warning prefix 정리
- [x] 3.3 Streaming 임계값 불일치: 코드와 README 설명을 현재 동작에 맞게 정리

### 리뷰 메모

- 전체 테스트 스위트는 JDK 17로 실행해 통과함
- Gradle 8.14.2는 현재 로컬 기본 JDK 25로 직접 실행하면 build script 로딩 단계에서 실패할 수 있음. 프로젝트 검증은 JDK 17 런타임으로 수행함

---

## Phase 1: Critical 이슈 수정

### 1.1 Java 버전 호환성 문제

**파일**: `build.gradle`

**현재 코드**:
```groovy
java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(23)
    }
    sourceCompatibility = JavaVersion.VERSION_1_8
    targetCompatibility = JavaVersion.VERSION_1_8
}
```

**문제점**:
- Java 23 toolchain으로 Java 8 바이트코드 생성 시도
- 실제 Java 8 환경에서 호환성 문제 발생 가능
- 일부 API가 Java 8에서 동작하지 않을 수 있음

**수정 방안**:
```groovy
java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(8)
    }
}

// 또는 더 넓은 호환성을 위해
java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(11)  // LTS 버전
    }
    sourceCompatibility = JavaVersion.VERSION_1_8
    targetCompatibility = JavaVersion.VERSION_1_8
}
```

**검증 방법**:
- [ ] Java 8 환경에서 빌드 테스트
- [ ] Java 8 환경에서 전체 테스트 스위트 실행
- [ ] Java 11, 17, 21 환경에서도 호환성 테스트

---

### 1.2 ChunkIterator 스레드 리소스 누수

**파일**: `src/main/java/com/excelninja/infrastructure/io/StreamingWorkbookReader.java`

**현재 코드**:
```java
@Override
public void close() {
    producerThread.interrupt();
    if (closeOnFinish && managedInputStream != null) {
        try {
            managedInputStream.close();
        } catch (IOException e) {
            logger.log(Level.WARNING, "Error closing input stream", e);
        }
    }
}
```

**문제점**:
- `interrupt()` 호출 후 스레드 종료 대기 없음
- 좀비 스레드 발생 가능
- 리소스 정리 순서 불명확

**수정 방안**:
```java
@Override
public void close() {
    // 1. 생산자 스레드에 종료 신호
    producerThread.interrupt();
    
    // 2. 스레드 종료 대기 (타임아웃 설정)
    try {
        producerThread.join(5000);  // 최대 5초 대기
        if (producerThread.isAlive()) {
            logger.warning("Producer thread did not terminate within timeout");
        }
    } catch (InterruptedException e) {
        Thread.currentThread().interrupt();
        logger.warning("Interrupted while waiting for producer thread");
    }
    
    // 3. 큐 정리
    queue.clear();
    
    // 4. InputStream 정리
    if (closeOnFinish && managedInputStream != null) {
        try {
            managedInputStream.close();
        } catch (IOException e) {
            logger.log(Level.WARNING, "Error closing input stream", e);
        }
    }
}
```

**추가 개선**:
```java
// AutoCloseable 명시적 구현 및 try-with-resources 지원
private class ChunkIterator<T> implements Iterator<List<T>>, AutoCloseable {
    
    private volatile boolean closed = false;
    
    @Override
    public boolean hasNext() {
        if (closed) {
            return false;
        }
        // ... 기존 로직
    }
    
    // finalizer 대신 Cleaner 사용 (Java 9+) 또는 명시적 close 권장
}
```

**검증 방법**:
- [ ] 스레드 덤프로 좀비 스레드 확인
- [ ] 반복적인 청크 읽기 후 스레드 수 확인
- [ ] `close()` 호출 후 리소스 해제 확인

---

### 1.3 EntityMetadata 캐시 메모리 누수

**파일**: `src/main/java/com/excelninja/infrastructure/metadata/EntityMetadata.java`

**현재 코드**:
```java
private static final Map<Class<?>, EntityMetadata<?>> METADATA_CACHE = new ConcurrentHashMap<>();
```

**문제점**:
- 캐시 크기 제한 없음
- eviction 정책 없음
- 동적 클래스 로딩 환경에서 메모리 누수

**수정 방안 A - WeakReference 사용**:
```java
import java.lang.ref.WeakReference;
import java.util.WeakHashMap;
import java.util.Collections;

private static final Map<Class<?>, WeakReference<EntityMetadata<?>>> METADATA_CACHE = 
    Collections.synchronizedMap(new WeakHashMap<>());

@SuppressWarnings("unchecked")
public static <T> EntityMetadata<T> of(Class<T> entityType) {
    WeakReference<EntityMetadata<?>> ref = METADATA_CACHE.get(entityType);
    EntityMetadata<?> cached = (ref != null) ? ref.get() : null;
    
    if (cached != null) {
        return (EntityMetadata<T>) cached;
    }
    
    EntityMetadata<T> newMetadata = new EntityMetadata<>(entityType);
    METADATA_CACHE.put(entityType, new WeakReference<>(newMetadata));
    return newMetadata;
}
```

**수정 방안 B - LRU 캐시 사용 (권장)**:
```java
import java.util.LinkedHashMap;
import java.util.Map;

private static final int MAX_CACHE_SIZE = 1000;

private static final Map<Class<?>, EntityMetadata<?>> METADATA_CACHE = 
    Collections.synchronizedMap(new LinkedHashMap<Class<?>, EntityMetadata<?>>(16, 0.75f, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry<Class<?>, EntityMetadata<?>> eldest) {
            boolean shouldRemove = size() > MAX_CACHE_SIZE;
            if (shouldRemove) {
                logger.fine("Evicting metadata cache entry for: " + eldest.getKey().getName());
            }
            return shouldRemove;
        }
    });
```

**수정 방안 C - Caffeine 캐시 사용 (외부 의존성)**:
```java
// build.gradle에 추가
// implementation 'com.github.ben-manes.caffeine:caffeine:3.1.8'

import com.github.benmanes.caffeine.cache.Cache;
import com.github.benmanes.caffeine.cache.Caffeine;

private static final Cache<Class<?>, EntityMetadata<?>> METADATA_CACHE = Caffeine.newBuilder()
    .maximumSize(1000)
    .expireAfterAccess(Duration.ofHours(1))
    .recordStats()
    .build();

@SuppressWarnings("unchecked")
public static <T> EntityMetadata<T> of(Class<T> entityType) {
    return (EntityMetadata<T>) METADATA_CACHE.get(entityType, EntityMetadata::new);
}
```

**검증 방법**:
- [ ] 캐시 크기 제한 동작 확인
- [ ] 메모리 프로파일링으로 누수 없음 확인
- [ ] 높은 동시성 환경에서 캐시 동작 확인

---

## Phase 2: High 이슈 수정

### 2.1 NinjaExcel Thread Safety

**파일**: `src/main/java/com/excelninja/application/facade/NinjaExcel.java`

**현재 코드**:
```java
private static final PoiWorkbookReader POI_WORKBOOK_READER = new PoiWorkbookReader();
private static final StreamingWorkbookReader STREAMING_WORKBOOK_READER = new StreamingWorkbookReader();
private static final PoiWorkbookWriter WORKBOOK_WRITER = new PoiWorkbookWriter();
private static final DefaultConverter CONVERTER = new DefaultConverter();
```

**문제점**:
- Reader/Writer 인스턴스가 stateless인지 확인 필요
- 멀티스레드 환경에서 공유 시 문제 가능

**수정 방안**:
```java
// 옵션 1: 매번 새 인스턴스 생성 (안전하지만 오버헤드)
private static WorkbookReader getReader(boolean useStreaming) {
    return useStreaming ? new StreamingWorkbookReader() : new PoiWorkbookReader();
}

// 옵션 2: Reader/Writer가 stateless임을 확인하고 문서화
/**
 * Thread-safe: Reader implementations are stateless and can be shared across threads.
 */
private static final PoiWorkbookReader POI_WORKBOOK_READER = new PoiWorkbookReader();

// 옵션 3: ThreadLocal 사용
private static final ThreadLocal<PoiWorkbookReader> POI_READER_LOCAL = 
    ThreadLocal.withInitial(PoiWorkbookReader::new);
```

**검증 방법**:
- [ ] Reader/Writer 클래스의 상태 변수 확인
- [ ] 멀티스레드 동시 읽기/쓰기 테스트
- [ ] 스트레스 테스트 추가

---

### 2.2 헤더 셀 타입 예외 처리

**파일**: `src/main/java/com/excelninja/infrastructure/io/PoiWorkbookReader.java`

**현재 코드**:
```java
for (Cell cell : headerRow) {
    String headerValue = cell.getStringCellValue();  // 숫자 셀이면 예외!
    // ...
}
```

**수정 방안**:
```java
for (Cell cell : headerRow) {
    String headerValue = getCellValueAsString(cell);
    if (headerValue == null || headerValue.trim().isEmpty()) {
        throw new InvalidDocumentStructureException(
            "Header cannot be empty at column " + cell.getColumnIndex() + " in sheet " + sheetName);
    }
    headerTitles.add(headerValue.trim());
}

private String getCellValueAsString(Cell cell) {
    if (cell == null) {
        return null;
    }
    
    switch (cell.getCellType()) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            // 숫자를 문자열로 변환 (정수면 소수점 제거)
            double numValue = cell.getNumericCellValue();
            if (numValue == Math.floor(numValue)) {
                return String.valueOf((long) numValue);
            }
            return String.valueOf(numValue);
        case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
        case FORMULA:
            try {
                return cell.getStringCellValue();
            } catch (IllegalStateException e) {
                return String.valueOf(cell.getNumericCellValue());
            }
        case BLANK:
            return null;
        default:
            return cell.toString();
    }
}
```

**검증 방법**:
- [ ] 숫자 헤더가 있는 Excel 파일 테스트
- [ ] 수식 헤더가 있는 Excel 파일 테스트
- [ ] 혼합 타입 헤더 테스트

---

### 2.3 OPCPackage InputStream 소비 문제

**파일**: `src/main/java/com/excelninja/infrastructure/io/StreamingWorkbookReader.java`

**현재 코드**:
```java
@Override
public ExcelWorkbook read(InputStream inputStream) throws IOException {
    try (OPCPackage opcPackage = OPCPackage.open(inputStream)) {
        return readFromOPCPackage(opcPackage);
    }
}
```

**문제점**:
- `OPCPackage.open(inputStream)`이 스트림을 완전히 소비
- 호출자가 스트림 재사용 불가능
- 문서화 부재

**수정 방안**:
```java
/**
 * Reads an Excel workbook from the given InputStream.
 * 
 * <p><b>Warning:</b> This method will fully consume the InputStream. 
 * The stream should not be reused after calling this method.
 * The caller is responsible for closing the InputStream after this method returns.
 *
 * @param inputStream the input stream to read from (will be consumed but not closed)
 * @return the parsed ExcelWorkbook
 * @throws IOException if an I/O error occurs
 */
@Override
public ExcelWorkbook read(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "InputStream cannot be null");
    
    try {
        // OPCPackage.open()은 내부적으로 스트림을 임시 파일로 복사할 수 있음
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        try {
            return readFromOPCPackage(opcPackage);
        } finally {
            opcPackage.close();
        }
    } catch (InvalidFormatException e) {
        throw new DocumentConversionException("Invalid Excel format", e);
    } catch (Exception e) {
        throw new DocumentConversionException("Failed to read Excel file with streaming reader", e);
    }
}
```

**검증 방법**:
- [ ] Javadoc 문서화 확인
- [ ] 스트림 재사용 시도 테스트 (예외 또는 빈 결과 확인)

---

## Phase 3: Medium 이슈 수정

### 3.1 README 버전 불일치

**파일**: `README.md`

**현재 코드**:
```groovy
implementation 'io.github.excel-ninja:excelNinja:0.0.8'
```

**수정 방안**:
```groovy
implementation 'io.github.excel-ninja:excelNinja:1.0.0'
```

---

### 3.2 로깅 레벨 일관성

**파일**: `NinjaExcel.java` 및 기타

**수정 방안 - 로깅 정책 정의**:

| 레벨 | 용도 |
|------|------|
| `SEVERE` | 복구 불가능한 오류 |
| `WARNING` | 복구 가능한 오류, 성능 저하 |
| `INFO` | 중요한 비즈니스 이벤트 (파일 읽기/쓰기 완료) |
| `FINE` | 디버깅용 상세 정보 |
| `FINER` | 메서드 진입/종료 |
| `FINEST` | 매우 상세한 디버깅 |

```java
// 변경 전
logger.info(String.format("[NINJA-EXCEL] Reading Excel file: %s", fileName));
logger.fine(String.format("[NINJA-EXCEL] Writing Excel workbook..."));

// 변경 후 - 일관된 INFO 레벨 사용
logger.info(String.format("[NINJA-EXCEL] Reading Excel file: %s (%.2f KB)", 
    fileName, fileSize / 1024.0));
logger.info(String.format("[NINJA-EXCEL] Writing Excel workbook: %d sheets, %d records", 
    sheetCount, recordCount));
```

---

### 3.3 Streaming 임계값 불일치

**파일**: `NinjaExcel.java`

**현재 코드**:
```java
private static final long STREAMING_THRESHOLD_BYTES = 10 * 1024 * 1024; // 10MB

// 하지만 read()에서는...
boolean useStreaming = fileSize > 1024 * 1024; // 1MB 하드코딩
```

**수정 방안**:
```java
private static final long STREAMING_THRESHOLD_BYTES = 10 * 1024 * 1024; // 10MB
private static volatile long streamingThreshold = STREAMING_THRESHOLD_BYTES;

public static void setStreamingThreshold(long thresholdBytes) {
    if (thresholdBytes <= 0) {
        throw new IllegalArgumentException("Threshold must be positive");
    }
    streamingThreshold = thresholdBytes;
    logger.info(String.format("[NINJA-EXCEL] Streaming threshold updated to %.2f MB",
            thresholdBytes / (1024.0 * 1024.0)));
}

private static boolean shouldUseStreaming(long fileSize) {
    return fileSize > streamingThreshold;
}

// read() 메서드에서
public static <T> List<T> read(File file, Class<T> clazz) {
    // ...
    boolean useStreaming = shouldUseStreaming(file.length());  // 일관된 임계값 사용
    // ...
}
```

---

## 테스트 계획

### 단위 테스트 추가

```java
// 1. 캐시 eviction 테스트
@Test
@DisplayName("EntityMetadata 캐시가 최대 크기를 초과하지 않음")
void metadataCacheShouldNotExceedMaxSize() {
    // MAX_CACHE_SIZE + 100개의 서로 다른 클래스 메타데이터 생성
    // 캐시 크기가 MAX_CACHE_SIZE 이하인지 확인
}

// 2. 스레드 안전성 테스트
@Test
@DisplayName("멀티스레드 환경에서 동시 읽기 안전")
void shouldBeThreadSafeForConcurrentReads() {
    ExecutorService executor = Executors.newFixedThreadPool(10);
    // 동시에 같은 파일 읽기
    // 모든 결과가 동일한지 확인
}

// 3. 리소스 누수 테스트
@Test
@DisplayName("ChunkIterator close 후 스레드 종료 확인")
void shouldTerminateThreadAfterClose() {
    Iterator<List<Employee>> chunks = NinjaExcel.readInChunks(file, Employee.class, 100);
    ((AutoCloseable) chunks).close();
    
    Thread.sleep(6000);  // join 타임아웃 + 여유
    // 활성 스레드 수 확인
}
```

### 통합 테스트

- [ ] 대용량 파일 (100MB+) 스트리밍 테스트
- [ ] 다양한 Excel 버전 호환성 테스트 (.xls, .xlsx, .xlsm)
- [ ] 멀티스레드 환경 부하 테스트

### 성능 테스트

- [ ] 캐시 히트율 측정
- [ ] 메모리 사용량 프로파일링
- [ ] 처리량 벤치마크 (records/sec)

---

## 릴리즈 체크리스트

### v1.0.1 패치 릴리즈

- [ ] Phase 1 (Critical) 이슈 모두 수정
- [ ] 관련 테스트 추가 및 통과
- [ ] CHANGELOG.md 업데이트
- [ ] README.md 버전 수정

### v1.1.0 마이너 릴리즈

- [ ] Phase 2 (High) 이슈 모두 수정
- [ ] Phase 3 (Medium) 이슈 모두 수정
- [ ] 전체 테스트 스위트 통과
- [ ] 성능 벤치마크 결과 문서화
- [ ] API 문서 (Javadoc) 업데이트

---

## 참고 자료

- [Apache POI Documentation](https://poi.apache.org/)
- [Java Concurrency in Practice](https://jcip.net/)
- [Effective Java 3rd Edition - Item 7: Eliminate obsolete object references](https://www.oreilly.com/library/view/effective-java-3rd/9780134686097/)

---

**문서 작성**: Claude  
**최종 수정**: 2025-01-11
