package com.example.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * One ordered item of document body content.
 *
 * <p>Deferred rather than a finished element, because an image is a package part plus a
 * relationship and no package exists until {@code build()}. Text implementations ignore
 * the argument.
 *
 * <p>Returns {@code Object} because an OOXML body holds a heterogeneous sequence:
 * {@code w:p} and {@code w:tbl} are both valid children, and docx4j models that as
 * {@code List<Object>}. The returned value must be an {@code org.docx4j.wml.P} or an
 * {@code org.docx4j.wml.Tbl}.
 *
 * <p>It lives here rather than in {@code content} because it names
 * {@code WordprocessingMLPackage}, which that package is kept free of.
 */
@FunctionalInterface
public interface DocumentContent {
    Object toBodyElement(WordprocessingMLPackage pkg);
}
