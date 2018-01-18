// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.generated;

import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.models.extensions.*;
import com.microsoft.graph.models.generated.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.requests.extensions.*;
import com.microsoft.graph.requests.generated.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;

import java.util.Arrays;
import java.util.EnumSet;
// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Base Workbook Chart Add Request Builder.
 */
public class BaseWorkbookChartAddRequestBuilder extends BaseActionRequestBuilder {

    /**
     * The request builder for this WorkbookChartAdd
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseWorkbookChartAddRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions, final String type, final com.google.gson.JsonElement sourceData, final String seriesBy) {
        super(requestUrl, client, requestOptions);
        bodyParams.put("type", type);
        bodyParams.put("sourceData", sourceData);
        bodyParams.put("seriesBy", seriesBy);
    }

    /**
     * Creates the IWorkbookChartAddRequest
     *
     * @return The IWorkbookChartAddRequest instance
     */
    public IWorkbookChartAddRequest buildRequest() {
        return buildRequest(getOptions());
    }

    /**
     * Creates the IWorkbookChartAddRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return The IWorkbookChartAddRequest instance
     */
    public IWorkbookChartAddRequest buildRequest(final java.util.List<? extends Option> requestOptions) {
        WorkbookChartAddRequest request = new WorkbookChartAddRequest(
                getRequestUrl(),
                getClient(),
                requestOptions
        );

        if (hasParameter("type")) {
            request.body.type = getParameter("type");
        }

        if (hasParameter("sourceData")) {
            request.body.sourceData = getParameter("sourceData");
        }

        if (hasParameter("seriesBy")) {
            request.body.seriesBy = getParameter("seriesBy");
        }

        return request;
    }
}
