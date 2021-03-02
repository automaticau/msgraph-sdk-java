// Template Source: BaseEntityRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.OrganizationalBranding;
import com.microsoft.graph.requests.extensions.IOrganizationalBrandingLocalizationCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IOrganizationalBrandingLocalizationRequestBuilder;
import com.microsoft.graph.requests.extensions.OrganizationalBrandingLocalizationCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.OrganizationalBrandingLocalizationRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequestBuilder;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Organizational Branding Request Builder.
 */
public class OrganizationalBrandingRequestBuilder extends BaseRequestBuilder implements IOrganizationalBrandingRequestBuilder {

    /**
     * The request builder for the OrganizationalBranding
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public OrganizationalBrandingRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the IOrganizationalBrandingRequest instance
     */
    public IOrganizationalBrandingRequest buildRequest(final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for this request
     * @return the IOrganizationalBrandingRequest instance
     */
    public IOrganizationalBrandingRequest buildRequest(final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        return new com.microsoft.graph.requests.extensions.OrganizationalBrandingRequest(getRequestUrl(), getClient(), requestOptions);
    }


    public IOrganizationalBrandingLocalizationCollectionRequestBuilder localizations() {
        return new OrganizationalBrandingLocalizationCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("localizations"), getClient(), null);
    }

    public IOrganizationalBrandingLocalizationRequestBuilder localizations(final String id) {
        return new OrganizationalBrandingLocalizationRequestBuilder(getRequestUrlWithAdditionalSegment("localizations") + "/" + id, getClient(), null);
    }
}