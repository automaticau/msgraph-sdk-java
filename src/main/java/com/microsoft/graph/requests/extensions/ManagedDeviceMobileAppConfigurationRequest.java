// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.ManagedDeviceMobileAppConfiguration;
import com.microsoft.graph.models.extensions.ManagedDeviceMobileAppConfigurationAssignment;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationAssignmentRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationAssignmentRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationDeviceStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationDeviceStatusRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationDeviceStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationDeviceStatusRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationDeviceSummaryRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationDeviceSummaryRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationUserStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationUserStatusRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationUserStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationUserStatusRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedDeviceMobileAppConfigurationUserSummaryRequestBuilder;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationUserSummaryRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed Device Mobile App Configuration Request.
 */
public class ManagedDeviceMobileAppConfigurationRequest extends BaseRequest implements IManagedDeviceMobileAppConfigurationRequest {
	
    /**
     * The request for the ManagedDeviceMobileAppConfiguration
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param responseClass  the class of the response
     */
    public ManagedDeviceMobileAppConfigurationRequest(final String requestUrl,
            final IBaseClient client,
            final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions,
            final Class<? extends ManagedDeviceMobileAppConfiguration> responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * The request for the ManagedDeviceMobileAppConfiguration
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public ManagedDeviceMobileAppConfigurationRequest(final String requestUrl, final IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, ManagedDeviceMobileAppConfiguration.class);
    }

    /**
     * Gets the ManagedDeviceMobileAppConfiguration from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(final ICallback<? super ManagedDeviceMobileAppConfiguration> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the ManagedDeviceMobileAppConfiguration from the service
     *
     * @return the ManagedDeviceMobileAppConfiguration from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    public ManagedDeviceMobileAppConfiguration get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(final ICallback<? super ManagedDeviceMobileAppConfiguration> callback) {
        send(HttpMethod.DELETE, callback, null);
    }

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {
        send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this ManagedDeviceMobileAppConfiguration with a source
     *
     * @param sourceManagedDeviceMobileAppConfiguration the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(final ManagedDeviceMobileAppConfiguration sourceManagedDeviceMobileAppConfiguration, final ICallback<? super ManagedDeviceMobileAppConfiguration> callback) {
        send(HttpMethod.PATCH, callback, sourceManagedDeviceMobileAppConfiguration);
    }

    /**
     * Patches this ManagedDeviceMobileAppConfiguration with a source
     *
     * @param sourceManagedDeviceMobileAppConfiguration the source object with updates
     * @return the updated ManagedDeviceMobileAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    public ManagedDeviceMobileAppConfiguration patch(final ManagedDeviceMobileAppConfiguration sourceManagedDeviceMobileAppConfiguration) throws ClientException {
        return send(HttpMethod.PATCH, sourceManagedDeviceMobileAppConfiguration);
    }

    /**
     * Creates a ManagedDeviceMobileAppConfiguration with a new object
     *
     * @param newManagedDeviceMobileAppConfiguration the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(final ManagedDeviceMobileAppConfiguration newManagedDeviceMobileAppConfiguration, final ICallback<? super ManagedDeviceMobileAppConfiguration> callback) {
        send(HttpMethod.POST, callback, newManagedDeviceMobileAppConfiguration);
    }

    /**
     * Creates a ManagedDeviceMobileAppConfiguration with a new object
     *
     * @param newManagedDeviceMobileAppConfiguration the new object to create
     * @return the created ManagedDeviceMobileAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    public ManagedDeviceMobileAppConfiguration post(final ManagedDeviceMobileAppConfiguration newManagedDeviceMobileAppConfiguration) throws ClientException {
        return send(HttpMethod.POST, newManagedDeviceMobileAppConfiguration);
    }

    /**
     * Creates a ManagedDeviceMobileAppConfiguration with a new object
     *
     * @param newManagedDeviceMobileAppConfiguration the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(final ManagedDeviceMobileAppConfiguration newManagedDeviceMobileAppConfiguration, final ICallback<? super ManagedDeviceMobileAppConfiguration> callback) {
        send(HttpMethod.PUT, callback, newManagedDeviceMobileAppConfiguration);
    }

    /**
     * Creates a ManagedDeviceMobileAppConfiguration with a new object
     *
     * @param newManagedDeviceMobileAppConfiguration the object to create/update
     * @return the created ManagedDeviceMobileAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    public ManagedDeviceMobileAppConfiguration put(final ManagedDeviceMobileAppConfiguration newManagedDeviceMobileAppConfiguration) throws ClientException {
        return send(HttpMethod.PUT, newManagedDeviceMobileAppConfiguration);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     public IManagedDeviceMobileAppConfigurationRequest select(final String value) {
         getQueryOptions().add(new com.microsoft.graph.options.QueryOption("$select", value));
         return (ManagedDeviceMobileAppConfigurationRequest)this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
     public IManagedDeviceMobileAppConfigurationRequest expand(final String value) {
         getQueryOptions().add(new com.microsoft.graph.options.QueryOption("$expand", value));
         return (ManagedDeviceMobileAppConfigurationRequest)this;
     }

}

