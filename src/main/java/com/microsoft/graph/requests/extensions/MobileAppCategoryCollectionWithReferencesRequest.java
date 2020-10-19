// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.MobileApp;
import com.microsoft.graph.models.extensions.MobileAppCategory;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.models.extensions.MobileAppCategory;
import com.microsoft.graph.requests.extensions.MobileAppCategoryWithReferenceRequest;
import com.microsoft.graph.requests.extensions.MobileAppCategoryReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.MobileAppCategoryWithReferenceRequestBuilder;
import com.microsoft.graph.requests.extensions.MobileAppCategoryCollectionWithReferencesRequest;
import com.microsoft.graph.requests.extensions.MobileAppCategoryCollectionWithReferencesRequest;
import com.microsoft.graph.requests.extensions.MobileAppCategoryCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseCollectionWithReferencesRequest;
import com.microsoft.graph.http.BaseCollectionWithReferencesRequestBuilder;
import com.microsoft.graph.concurrency.IExecutors;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Mobile App Category Collection With References Request.
 */
public class MobileAppCategoryCollectionWithReferencesRequest extends BaseCollectionWithReferencesRequest<MobileAppCategory, MobileAppCategoryWithReferenceRequest, MobileAppCategoryReferenceRequestBuilder, MobileAppCategoryWithReferenceRequestBuilder, MobileAppCategoryCollectionResponse, MobileAppCategoryCollectionPage, MobileAppCategoryCollectionWithReferencesRequest> {

    /**
     * The request builder for this collection of MobileAppCategory
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public MobileAppCategoryCollectionWithReferencesRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, MobileAppCategoryCollectionResponse.class, MobileAppCategoryCollectionPage.class, MobileAppCategoryCollectionWithReferencesRequestBuilder.class);
    }

    public MobileAppCategoryCollectionWithReferencesRequest expand(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$expand", value));
        return this;
    }

    public MobileAppCategoryCollectionWithReferencesRequest filter(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$filter", value));
        return this;
    }

    public MobileAppCategoryCollectionWithReferencesRequest orderBy(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$orderby", value));
        return this;
    }

    public MobileAppCategoryCollectionWithReferencesRequest select(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$select", value));
        return this;
    }

    public MobileAppCategoryCollectionWithReferencesRequest top(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$top", value + ""));
        return this;
    }

}
