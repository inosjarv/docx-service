package com.example.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;

/**
 * One ordered item of document body content.
 *
 * <p>Deferred rather than a finished {@code P}, because an image is a package part plus
 * a relationship and no package exists until {@code build()}. Text implementations
 * ignore the argument.
 *
 * <p>It lives here rather than in {@code content} because it names
 * {@code WordprocessingMLPackage}, which that package is kept free of.
 */
@FunctionalInterface
public interface DocumentContent {
    P toParagraph(WordprocessingMLPackage pkg);
}
