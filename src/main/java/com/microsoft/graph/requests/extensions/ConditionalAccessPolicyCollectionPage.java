// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.ConditionalAccessPolicy;
import com.microsoft.graph.requests.extensions.ConditionalAccessPolicyCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.ConditionalAccessPolicyCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Conditional Access Policy Collection Page.
 */
public class ConditionalAccessPolicyCollectionPage extends BaseCollectionPage<ConditionalAccessPolicy, ConditionalAccessPolicyCollectionRequestBuilder> {

    /**
     * A collection page for ConditionalAccessPolicy
     *
     * @param response the serialized ConditionalAccessPolicyCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ConditionalAccessPolicyCollectionPage(@Nonnull final ConditionalAccessPolicyCollectionResponse response, @Nonnull final ConditionalAccessPolicyCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ConditionalAccessPolicy
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ConditionalAccessPolicyCollectionPage(@Nonnull final java.util.List<ConditionalAccessPolicy> pageContents, @Nullable final ConditionalAccessPolicyCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
