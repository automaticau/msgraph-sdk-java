// Template Source: BaseMethodCollectionRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.LocaleInfo;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

import com.microsoft.graph.requests.OutlookUserSupportedLanguagesCollectionRequestBuilder;
import com.microsoft.graph.requests.OutlookUserSupportedLanguagesCollectionRequest;
import com.microsoft.graph.requests.OutlookUserSupportedLanguagesCollectionResponse;
import com.microsoft.graph.options.FunctionOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseFunctionCollectionRequestBuilder;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Outlook User Supported Languages Collection Request Builder.
 */
public class OutlookUserSupportedLanguagesCollectionRequestBuilder extends BaseFunctionCollectionRequestBuilder<LocaleInfo, OutlookUserSupportedLanguagesCollectionRequestBuilder, OutlookUserSupportedLanguagesCollectionResponse, OutlookUserSupportedLanguagesCollectionPage, OutlookUserSupportedLanguagesCollectionRequest> {

    /**
     * The request builder for this collection of OutlookUser
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public OutlookUserSupportedLanguagesCollectionRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, OutlookUserSupportedLanguagesCollectionRequestBuilder.class, OutlookUserSupportedLanguagesCollectionRequest.class);
    }
    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the OutlookUserSupportedLanguagesCollectionRequest instance
     */
    @Override
    @Nonnull
    public OutlookUserSupportedLanguagesCollectionRequest buildRequest(@Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        final OutlookUserSupportedLanguagesCollectionRequest request = super.buildRequest(requestOptions);
        return request;
    }
}
