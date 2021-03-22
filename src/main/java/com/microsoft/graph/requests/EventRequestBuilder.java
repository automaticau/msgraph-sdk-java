// Template Source: BaseEntityRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.Event;
import com.microsoft.graph.models.Recipient;
import com.microsoft.graph.models.DateTimeTimeZone;
import com.microsoft.graph.requests.AttachmentCollectionRequestBuilder;
import com.microsoft.graph.requests.AttachmentRequestBuilder;
import com.microsoft.graph.requests.CalendarRequestBuilder;
import com.microsoft.graph.requests.ExtensionCollectionRequestBuilder;
import com.microsoft.graph.requests.ExtensionRequestBuilder;
import com.microsoft.graph.requests.EventCollectionRequestBuilder;
import com.microsoft.graph.requests.EventRequestBuilder;
import com.microsoft.graph.requests.MultiValueLegacyExtendedPropertyCollectionRequestBuilder;
import com.microsoft.graph.requests.MultiValueLegacyExtendedPropertyRequestBuilder;
import com.microsoft.graph.requests.SingleValueLegacyExtendedPropertyCollectionRequestBuilder;
import com.microsoft.graph.requests.SingleValueLegacyExtendedPropertyRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequestBuilder;
import com.microsoft.graph.models.EventAcceptParameterSet;
import com.microsoft.graph.models.EventCancelParameterSet;
import com.microsoft.graph.models.EventDeclineParameterSet;
import com.microsoft.graph.models.EventForwardParameterSet;
import com.microsoft.graph.models.EventSnoozeReminderParameterSet;
import com.microsoft.graph.models.EventTentativelyAcceptParameterSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Event Request Builder.
 */
public class EventRequestBuilder extends BaseRequestBuilder<Event> {

    /**
     * The request builder for the Event
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public EventRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the EventRequest instance
     */
    @Nonnull
    public EventRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for this request
     * @return the EventRequest instance
     */
    @Nonnull
    public EventRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        return new com.microsoft.graph.requests.EventRequest(getRequestUrl(), getClient(), requestOptions);
    }


    /**
     *  Gets a request builder for the Attachment collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public AttachmentCollectionRequestBuilder attachments() {
        return new AttachmentCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("attachments"), getClient(), null);
    }

    /**
     * Gets a request builder for the Attachment item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public AttachmentRequestBuilder attachments(@Nonnull final String id) {
        return new AttachmentRequestBuilder(getRequestUrlWithAdditionalSegment("attachments") + "/" + id, getClient(), null);
    }

    /**
     * Gets the request builder for Calendar
     *
     * @return the CalendarRequestBuilder instance
     */
    @Nonnull
    public CalendarRequestBuilder calendar() {
        return new CalendarRequestBuilder(getRequestUrlWithAdditionalSegment("calendar"), getClient(), null);
    }
    /**
     *  Gets a request builder for the Extension collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public ExtensionCollectionRequestBuilder extensions() {
        return new ExtensionCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("extensions"), getClient(), null);
    }

    /**
     * Gets a request builder for the Extension item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public ExtensionRequestBuilder extensions(@Nonnull final String id) {
        return new ExtensionRequestBuilder(getRequestUrlWithAdditionalSegment("extensions") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the Event collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public EventCollectionRequestBuilder instances() {
        return new EventCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("instances"), getClient(), null);
    }

    /**
     * Gets a request builder for the Event item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public EventRequestBuilder instances(@Nonnull final String id) {
        return new EventRequestBuilder(getRequestUrlWithAdditionalSegment("instances") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the MultiValueLegacyExtendedProperty collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public MultiValueLegacyExtendedPropertyCollectionRequestBuilder multiValueExtendedProperties() {
        return new MultiValueLegacyExtendedPropertyCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("multiValueExtendedProperties"), getClient(), null);
    }

    /**
     * Gets a request builder for the MultiValueLegacyExtendedProperty item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public MultiValueLegacyExtendedPropertyRequestBuilder multiValueExtendedProperties(@Nonnull final String id) {
        return new MultiValueLegacyExtendedPropertyRequestBuilder(getRequestUrlWithAdditionalSegment("multiValueExtendedProperties") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the SingleValueLegacyExtendedProperty collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public SingleValueLegacyExtendedPropertyCollectionRequestBuilder singleValueExtendedProperties() {
        return new SingleValueLegacyExtendedPropertyCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("singleValueExtendedProperties"), getClient(), null);
    }

    /**
     * Gets a request builder for the SingleValueLegacyExtendedProperty item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public SingleValueLegacyExtendedPropertyRequestBuilder singleValueExtendedProperties(@Nonnull final String id) {
        return new SingleValueLegacyExtendedPropertyRequestBuilder(getRequestUrlWithAdditionalSegment("singleValueExtendedProperties") + "/" + id, getClient(), null);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventAcceptRequestBuilder accept(@Nonnull final EventAcceptParameterSet parameters) {
        return new EventAcceptRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.accept"), getClient(), null, parameters);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventCancelRequestBuilder cancel(@Nonnull final EventCancelParameterSet parameters) {
        return new EventCancelRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.cancel"), getClient(), null, parameters);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventDeclineRequestBuilder decline(@Nonnull final EventDeclineParameterSet parameters) {
        return new EventDeclineRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.decline"), getClient(), null, parameters);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     */
    @Nonnull
    public EventDismissReminderRequestBuilder dismissReminder() {
        return new EventDismissReminderRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.dismissReminder"), getClient(), null);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventForwardRequestBuilder forward(@Nonnull final EventForwardParameterSet parameters) {
        return new EventForwardRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.forward"), getClient(), null, parameters);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventSnoozeReminderRequestBuilder snoozeReminder(@Nonnull final EventSnoozeReminderParameterSet parameters) {
        return new EventSnoozeReminderRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.snoozeReminder"), getClient(), null, parameters);
    }

    /**
     * Gets a builder to execute the method
     * @return the request builder
     * @param parameters the parameters for the service method
     */
    @Nonnull
    public EventTentativelyAcceptRequestBuilder tentativelyAccept(@Nonnull final EventTentativelyAcceptParameterSet parameters) {
        return new EventTentativelyAcceptRequestBuilder(getRequestUrlWithAdditionalSegment("microsoft.graph.tentativelyAccept"), getClient(), null, parameters);
    }
}