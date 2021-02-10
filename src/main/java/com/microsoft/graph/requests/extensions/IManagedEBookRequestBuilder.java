// Template Source: IBaseEntityRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.ManagedEBook;
import com.microsoft.graph.models.extensions.ManagedEBookAssignment;
import com.microsoft.graph.requests.extensions.IManagedEBookAssignmentCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedEBookAssignmentRequestBuilder;
import com.microsoft.graph.requests.extensions.IDeviceInstallStateCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IDeviceInstallStateRequestBuilder;
import com.microsoft.graph.requests.extensions.IEBookInstallSummaryRequestBuilder;
import com.microsoft.graph.requests.extensions.IUserInstallStateSummaryCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IUserInstallStateSummaryRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The interface for the Managed EBook Request Builder.
 */
public interface IManagedEBookRequestBuilder extends IRequestBuilder {
    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the IManagedEBookRequest instance
     */
    IManagedEBookRequest buildRequest(final com.microsoft.graph.options.Option... requestOptions);

    /**
     * Creates the request with specific options instead of the existing options
     *
     * @param requestOptions the options for this request
     * @return the IManagedEBookRequest instance
     */
    IManagedEBookRequest buildRequest(final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions);


    IManagedEBookAssignmentCollectionRequestBuilder assignments();

    IManagedEBookAssignmentRequestBuilder assignments(final String id);

    IDeviceInstallStateCollectionRequestBuilder deviceStates();

    IDeviceInstallStateRequestBuilder deviceStates(final String id);

    /**
     * Gets the request builder for EBookInstallSummary
     *
     * @return the IEBookInstallSummaryRequestBuilder instance
     */
    IEBookInstallSummaryRequestBuilder installSummary();

    IUserInstallStateSummaryCollectionRequestBuilder userStateSummary();

    IUserInstallStateSummaryRequestBuilder userStateSummary(final String id);
    IManagedEBookAssignRequestBuilder assign(final java.util.List<ManagedEBookAssignment> managedEBookAssignments);

}