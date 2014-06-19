'use strict';

//Menu service used for managing  menus
angular.module('todoApp').service('todoService', ['loggingService',
	function(loggingService) {
		this.readAll = function($scope) {

			$scope.todos = [];
			$scope.open = [];
			$scope.inProgress = [];
			$scope.done = [];

			try {
				if (window.external !== undefined && window.external.OutlookApplication !== undefined) {
					var ol = window.external.OutlookApplication;

					var ns = ol.GetNameSpace("MAPI");

					var inbox = ns.GetDefaultFolder(28); //see http://msdn.microsoft.com/en-US/library/office/ff861868(v=office.15).aspx
					//to access sub folders, use .Folders(1)

					var items = inbox.Items;
					items.Sort("Importance", true);
					loggingService.info("Items count" + items.Count);
					
					for (i=0; i < items.Count; i++) {
						var item = items(i);

						if (item === undefined)
							break;

						loggingService.info(i + " " + item.Status + " " + item.Subject);

						if (item.Status == undefined) {
							// MailItem
							switch (item.FlagStatus) {
								// olNoFlag
								case 0:
									// $scope.open.push(createTodoFromOutlook(item));
									break;
									// olFlagComplete
								case 1:
									$scope.done.push(createTodoFromOutlook(item));
									break;
									// olFlagMarked
								case 2:
									$scope.open.push(createTodoFromOutlook(item));
									break;
								default:
									break;
							}
						} else {
							// TaskItem
							switch (item.Status) {
								case 0:
									$scope.open.push(createTodoFromOutlook(item));
									break;
								case 1:
									$scope.inProgress.push(createTodoFromOutlook(item));
									break;
								case 2:
									$scope.done.push(createTodoFromOutlook(item));
									break;
								default:
									break;
							}
						}
					}
				} else {
					$scope.outsideOfOutlook = true;
				}
			} catch (e) {
				loggingService.info(e);
			}

		}

		var createTodoFromOutlook = function(item) {
			return {
				text: item.Subject,
				status: item.Status,
				priority: item.Importance,
				notes: item.Notes
			};
		}
	}
]);
