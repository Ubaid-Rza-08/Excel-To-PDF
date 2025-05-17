FROM openjdk:24-ea-jdk-slim

WORKDIR /app

COPY target/excel_to_PDF-0.0.1-SNAPSHOT.jar app.jar

EXPOSE 8080

ENTRYPOINT ["java", "-jar", "app.jar"]
