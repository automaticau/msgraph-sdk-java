// Template Source: BaseMethodCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.ExchangeIdFormat;
import com.microsoft.graph.models.ConvertIdResult;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

import com.microsoft.graph.http.BaseCollectionPage;
import com.microsoft.graph.requests.UserTranslateExchangeIdsCollectionRequestBuilder;
import com.microsoft.graph.requests.UserTranslateExchangeIdsCollectionPage;
import com.microsoft.graph.requests.UserTranslateExchangeIdsCollectionResponse;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the User Translate Exchange Ids Collection Page.
 */
public class UserTranslateExchangeIdsCollectionPage extends BaseCollectionPage<ConvertIdResult, UserTranslateExchangeIdsCollectionRequestBuilder> {

    /**
     * A collection page for ConvertIdResult.
     *
     * @param response The serialized UserTranslateExchangeIdsCollectionResponse from the service
     * @param builder The request builder for the next collection page
     */
    public UserTranslateExchangeIdsCollectionPage(@Nonnull final UserTranslateExchangeIdsCollectionResponse response, @Nonnull final UserTranslateExchangeIdsCollectionRequestBuilder builder) {
       super(response, builder);
    }

	/**
     * Creates the collection page for UserTranslateExchangeIds
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public UserTranslateExchangeIdsCollectionPage(@Nonnull final java.util.List<ConvertIdResult> pageContents, @Nullable final UserTranslateExchangeIdsCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}