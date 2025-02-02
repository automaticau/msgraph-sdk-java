// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.WorkbookCommentReply;
import com.microsoft.graph.requests.WorkbookCommentReplyCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.WorkbookCommentReplyCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Workbook Comment Reply Collection Page.
 */
public class WorkbookCommentReplyCollectionPage extends BaseCollectionPage<WorkbookCommentReply, WorkbookCommentReplyCollectionRequestBuilder> {

    /**
     * A collection page for WorkbookCommentReply
     *
     * @param response the serialized WorkbookCommentReplyCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public WorkbookCommentReplyCollectionPage(@Nonnull final WorkbookCommentReplyCollectionResponse response, @Nonnull final WorkbookCommentReplyCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for WorkbookCommentReply
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public WorkbookCommentReplyCollectionPage(@Nonnull final java.util.List<WorkbookCommentReply> pageContents, @Nullable final WorkbookCommentReplyCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
