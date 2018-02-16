import requests
import sys
from datetime import date, datetime
#from dateutil.parser import parse


local_url = 'http://localhost:8081/report'


class ServiceManager:

    def __init__(self, from_date, to_date, vendor):
        """
        (str, str, str) -> ServiceManager
        Constructor of the class.
        Takes as arguments the search dates for the query.
        """
        self.__from_date = from_date
        self.__to_date = to_date
        self.__answers_submit_url = local_url

    def __make_request(self, url):
        """
        (str) -> dict
        Makes an http request in the specified url
        and returns the response decoded in json format
        """
        try:
            params = {"startDate": self.__from_date, "endDate": self.__to_date}
            response = requests.get(url, params=params)
            json_response = response.json()
            if (response.status_code != 200):
                raise Exception("code: {0}: reason: {1}".format(
                    response.status_code,
                    json_response)
                                )
            return json_response
        except Exception as e:
                print "Error while requesting data from server"
                print "{0}".format(e)
                sys.exit(1)

    def __format_data(self, json_data):
        """
        (None) -> list
        Gets the necessary data from the json file
        and returns them as a list of lists
        """
        data_list = []
        for reporting_aggregated_data in json_data:
            single_data_list = []
            print reporting_aggregated_data['dateCreated']
            single_data_list.append(datetime.strptime(reporting_aggregated_data['dayBegin'], '%Y-%m-%d'))
            single_data_list.append(datetime.strptime(reporting_aggregated_data['dayEnd'], '%Y-%m-%d'))
            single_data_list.append(datetime.fromtimestamp(reporting_aggregated_data['dateCreated'] // 1000))
            single_data_list.append(reporting_aggregated_data['tariff'])
            single_data_list.append(reporting_aggregated_data['ucgPopulation'])
            single_data_list.append(reporting_aggregated_data['utgPopulation'])
            single_data_list.append(reporting_aggregated_data['isAllowed'])
            single_data_list.append(reporting_aggregated_data['revenueSumUCG'])
            single_data_list.append(reporting_aggregated_data['revenueSumUTG'])
            single_data_list.append(reporting_aggregated_data['topupCountUCG'])
            single_data_list.append(reporting_aggregated_data['topupCountUTG'])
            single_data_list.append(reporting_aggregated_data['topupSumUCG'])
            single_data_list.append(reporting_aggregated_data['topupSumUTG'])
            single_data_list.append(reporting_aggregated_data['activeUTG'])
            single_data_list.append(reporting_aggregated_data['activeUCG'])
            single_data_list.append(reporting_aggregated_data[
                'uniqueInvitedUTG'
                ])
            single_data_list.append(reporting_aggregated_data[
                'uniqueInvitedUCG'
                ])
            single_data_list.append(reporting_aggregated_data[
                'uniqueRespondersUTG'
                ])

            data_list.append(single_data_list)

        return data_list

    def get_reporting_data(self):
        """
        (None) -> list
        Returns the reporting data
        """
        return self.__format_data(
            self.__make_request(
                self.__answers_submit_url
                )
            )


if __name__ == "__main__":

    service_manager = ServiceManager(
        "01/07/2016 00:00:00",
        "26/07/2016 23:59:59",
        "local"
        )
    print(service_manager.get_reporting_data())
