// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.RemoteAssistancePartner;
import com.microsoft.graph.requests.RemoteAssistancePartnerCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.RemoteAssistancePartnerCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Remote Assistance Partner Collection Page.
 */
public class RemoteAssistancePartnerCollectionPage extends BaseCollectionPage<RemoteAssistancePartner, RemoteAssistancePartnerCollectionRequestBuilder> {

    /**
     * A collection page for RemoteAssistancePartner
     *
     * @param response the serialized RemoteAssistancePartnerCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public RemoteAssistancePartnerCollectionPage(@Nonnull final RemoteAssistancePartnerCollectionResponse response, @Nonnull final RemoteAssistancePartnerCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for RemoteAssistancePartner
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public RemoteAssistancePartnerCollectionPage(@Nonnull final java.util.List<RemoteAssistancePartner> pageContents, @Nullable final RemoteAssistancePartnerCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
