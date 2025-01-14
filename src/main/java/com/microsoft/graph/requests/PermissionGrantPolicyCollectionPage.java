// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.PermissionGrantPolicy;
import com.microsoft.graph.requests.PermissionGrantPolicyCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.PermissionGrantPolicyCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Permission Grant Policy Collection Page.
 */
public class PermissionGrantPolicyCollectionPage extends BaseCollectionPage<PermissionGrantPolicy, PermissionGrantPolicyCollectionRequestBuilder> {

    /**
     * A collection page for PermissionGrantPolicy
     *
     * @param response the serialized PermissionGrantPolicyCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public PermissionGrantPolicyCollectionPage(@Nonnull final PermissionGrantPolicyCollectionResponse response, @Nonnull final PermissionGrantPolicyCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for PermissionGrantPolicy
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public PermissionGrantPolicyCollectionPage(@Nonnull final java.util.List<PermissionGrantPolicy> pageContents, @Nullable final PermissionGrantPolicyCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
