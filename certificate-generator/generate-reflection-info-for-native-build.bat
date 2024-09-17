REM This profiles the application and gathers reflection data using the native agent.
REM https://graalvm.github.io/native-build-tools/latest/gradle-plugin.html#agent-support

REM Runs on JVM with native-image-agent.
call gradlew -Pagent run

REM Copies the metadata collected by the agent into the project sources
./gradlew metadataCopy --task run --dir src/main/resources/META-INF/native-image

REM Builds image using metadata acquired by the agent.
REM ./gradlew nativeCompile
