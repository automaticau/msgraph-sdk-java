// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.EducationSubmission;
import com.microsoft.graph.requests.EducationSubmissionCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.EducationSubmissionCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Education Submission Collection Page.
 */
public class EducationSubmissionCollectionPage extends BaseCollectionPage<EducationSubmission, EducationSubmissionCollectionRequestBuilder> {

    /**
     * A collection page for EducationSubmission
     *
     * @param response the serialized EducationSubmissionCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public EducationSubmissionCollectionPage(@Nonnull final EducationSubmissionCollectionResponse response, @Nonnull final EducationSubmissionCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for EducationSubmission
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public EducationSubmissionCollectionPage(@Nonnull final java.util.List<EducationSubmission> pageContents, @Nullable final EducationSubmissionCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
