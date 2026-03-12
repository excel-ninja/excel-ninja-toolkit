package com.excelninja.infrastructure.metadata;

import com.excelninja.domain.annotation.ExcelReadColumn;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;
import java.io.IOException;
import java.net.URL;
import java.net.URLClassLoader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import static org.assertj.core.api.Assertions.assertThat;

@DisplayName("EntityMetadata cache")
class EntityMetadataCacheTest {

    @TempDir
    Path tempDir;

    static class AnnotatedDto {
        @ExcelReadColumn(headerName = "Value")
        private String value;

        public AnnotatedDto() {}
    }

    @AfterEach
    void resetCacheConfiguration() {
        EntityMetadata.clearCache();
        EntityMetadata.resetMaxCacheSizeForTesting();
    }

    @Test
    @DisplayName("EntityMetadata cache does not exceed configured max size")
    void metadataCacheShouldNotExceedMaxSize() throws Exception {
        EntityMetadata.clearCache();
        EntityMetadata.setMaxCacheSizeForTesting(5);

        try (URLClassLoader classLoader = compileGeneratedDtos(7)) {
            for (int i = 0; i < 7; i++) {
                Class<?> dtoClass = Class.forName("generated.metadata.GeneratedDto" + i, true, classLoader);
                EntityMetadata.of(dtoClass);
            }
        }

        assertThat(EntityMetadata.getCacheSize()).isEqualTo(EntityMetadata.getMaxCacheSize());
    }

    @Test
    @DisplayName("Concurrent lookups share one cached metadata instance")
    void shouldShareSingleMetadataInstanceAcrossThreads() throws Exception {
        EntityMetadata.clearCache();
        ExecutorService executorService = Executors.newFixedThreadPool(8);

        try {
            List<Callable<EntityMetadata<AnnotatedDto>>> tasks = new ArrayList<>();
            for (int i = 0; i < 20; i++) {
                tasks.add(() -> EntityMetadata.of(AnnotatedDto.class));
            }

            List<Future<EntityMetadata<AnnotatedDto>>> futures = executorService.invokeAll(tasks);
            EntityMetadata<AnnotatedDto> first = futures.get(0).get();

            for (Future<EntityMetadata<AnnotatedDto>> future : futures) {
                assertThat(future.get()).isSameAs(first);
            }
        } finally {
            executorService.shutdownNow();
        }

        assertThat(EntityMetadata.getCacheSize()).isEqualTo(1);
    }

    private URLClassLoader compileGeneratedDtos(int classCount) throws IOException {
        JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
        assertThat(compiler).isNotNull();

        Path sourceDir = tempDir.resolve("generated-sources");
        Path classesDir = tempDir.resolve("generated-classes");
        Files.createDirectories(sourceDir);
        Files.createDirectories(classesDir);

        List<Path> sourceFiles = new ArrayList<>();
        for (int i = 0; i < classCount; i++) {
            String className = "GeneratedDto" + i;
            Path sourceFile = sourceDir.resolve(className + ".java");
            Files.write(sourceFile, buildSource(className).getBytes(StandardCharsets.UTF_8));
            sourceFiles.add(sourceFile);
        }

        try (StandardJavaFileManager fileManager = compiler.getStandardFileManager(null, null, null)) {
            Iterable<? extends JavaFileObject> compilationUnits =
                    fileManager.getJavaFileObjectsFromFiles(toFiles(sourceFiles));

            List<String> options = Arrays.asList(
                    "-d", classesDir.toString(),
                    "-classpath", System.getProperty("java.class.path")
            );

            Boolean success = compiler.getTask(null, fileManager, null, options, null, compilationUnits).call();
            assertThat(success).isTrue();
        }

        return new URLClassLoader(new URL[]{classesDir.toUri().toURL()}, EntityMetadataCacheTest.class.getClassLoader());
    }

    private List<java.io.File> toFiles(List<Path> paths) {
        List<java.io.File> files = new ArrayList<>();
        for (Path path : paths) {
            files.add(path.toFile());
        }
        return files;
    }

    private String buildSource(String className) {
        return "package generated.metadata;\n"
                + "import com.excelninja.domain.annotation.ExcelReadColumn;\n"
                + "public class " + className + " {\n"
                + "    @ExcelReadColumn(headerName = \"Value\")\n"
                + "    private String value;\n"
                + "    public " + className + "() {}\n"
                + "}\n";
    }
}
