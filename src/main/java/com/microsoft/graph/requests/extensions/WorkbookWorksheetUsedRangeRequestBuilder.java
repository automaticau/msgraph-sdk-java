// Template Source: BaseMethodRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.requests.extensions.WorkbookWorksheetUsedRangeRequest;
import com.microsoft.graph.models.extensions.WorkbookWorksheet;
import com.microsoft.graph.models.extensions.WorkbookRange;
import com.microsoft.graph.models.extensions.WorkbookRange;
import com.microsoft.graph.http.BaseFunctionRequestBuilder;
import com.microsoft.graph.models.extensions.WorkbookWorksheetUsedRangeParameterSet;
import com.microsoft.graph.core.IBaseClient;
import com.google.gson.JsonElement;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Workbook Worksheet Used Range Request Builder.
 */
public class WorkbookWorksheetUsedRangeRequestBuilder extends BaseFunctionRequestBuilder<WorkbookRange> {

    /**
     * The request builder for this WorkbookWorksheetUsedRange
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public WorkbookWorksheetUsedRangeRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }
    /**
     * The request builder for this WorkbookWorksheetUsedRange
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param parameters     the parameters for the service method
     */
    public WorkbookWorksheetUsedRangeRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions, @Nonnull final WorkbookWorksheetUsedRangeParameterSet parameters) {
        super(requestUrl, client, requestOptions);
        if(parameters != null) {
            functionOptions = parameters.getFunctionOptions();
        }
    }

    /**
     * Creates the WorkbookWorksheetUsedRangeRequest
     *
     * @param requestOptions the options for the request
     * @return the WorkbookWorksheetUsedRangeRequest instance
     */
    @Nonnull
    public WorkbookWorksheetUsedRangeRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the WorkbookWorksheetUsedRangeRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return the WorkbookWorksheetUsedRangeRequest instance
     */
    @Nonnull
    public WorkbookWorksheetUsedRangeRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        WorkbookWorksheetUsedRangeRequest request = new WorkbookWorksheetUsedRangeRequest(
                getRequestUrl(),
                getClient(),
                requestOptions
        );

        return request;
    }

    /**
     * Gets the request builder for WorkbookRangeFormat
     *
     * @return the WorkbookRangeFormatRequestBuilder instance
     */
    @Nonnull
    public WorkbookRangeFormatRequestBuilder format() {
        return new WorkbookRangeFormatRequestBuilder(getRequestUrlWithAdditionalSegment("format"), getClient(), null);
    }

    /**
     * Gets the request builder for WorkbookRangeSort
     *
     * @return the WorkbookRangeSortRequestBuilder instance
     */
    @Nonnull
    public WorkbookRangeSortRequestBuilder sort() {
        return new WorkbookRangeSortRequestBuilder(getRequestUrlWithAdditionalSegment("sort"), getClient(), null);
    }

    /**
     * Gets the request builder for WorkbookWorksheet
     *
     * @return the WorkbookWorksheetRequestBuilder instance
     */
    @Nonnull
    public WorkbookWorksheetRequestBuilder worksheet() {
        return new WorkbookWorksheetRequestBuilder(getRequestUrlWithAdditionalSegment("worksheet"), getClient(), null);
    }
}
