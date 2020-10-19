// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.AdministrativeUnit;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.requests.extensions.AdministrativeUnitDeltaCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.AdministrativeUnitDeltaCollectionResponse;
import com.microsoft.graph.models.extensions.AdministrativeUnit;
import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseCollectionRequest;
import com.microsoft.graph.concurrency.IExecutors;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Administrative Unit Delta Collection Request.
 */
public class AdministrativeUnitDeltaCollectionRequest extends BaseCollectionRequest<AdministrativeUnit, AdministrativeUnitDeltaCollectionResponse, AdministrativeUnitDeltaCollectionPage> {


    /**
     * The request for this AdministrativeUnitDelta
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public AdministrativeUnitDeltaCollectionRequest(final String requestUrl, final IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, AdministrativeUnitDeltaCollectionResponse.class, AdministrativeUnitDeltaCollectionPage.class, AdministrativeUnitDeltaCollectionRequestBuilder.class);
    }


    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    public AdministrativeUnitDeltaCollectionRequest select(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$select", value));
        return this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value the max number of items to return
     * @return the updated request
     */
    public AdministrativeUnitDeltaCollectionRequest top(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$top", value+""));
        return this;
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    public AdministrativeUnitDeltaCollectionRequest expand(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$expand", value));
        return this;
    }

    /**
     * Sets the filter clause for the request
     *
     * @param value the filter clause
     * @return the updated request
     */
    public AdministrativeUnitDeltaCollectionRequest filter(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$filter", value));
        return this;
    }

    /**
     * Sets the order by clause for the request
     *
     * @param value the order by clause
     * @return the updated request
     */
    public AdministrativeUnitDeltaCollectionRequest orderBy(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$orderby", value));
        return this;
    }

}
