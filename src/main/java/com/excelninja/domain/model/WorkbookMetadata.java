package com.excelninja.domain.model;

import java.time.LocalDateTime;

public class WorkbookMetadata {
    private final String author;
    private final String title;
    private final LocalDateTime createdDate;

    public WorkbookMetadata() {
        this(null, null, LocalDateTime.now());
    }

    public WorkbookMetadata(
            String author,
            String title,
            LocalDateTime createdDate
    ) {
        this.author = author;
        this.title = title;
        this.createdDate = createdDate != null ? createdDate : LocalDateTime.now();
    }

    public String getAuthor() {
        return author;
    }

    public String getTitle() {
        return title;
    }

    public LocalDateTime getCreatedDate() {
        return createdDate;
    }
}