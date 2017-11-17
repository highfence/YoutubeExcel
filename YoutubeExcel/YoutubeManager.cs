using Google.Apis.Services;
using Google.Apis.YouTube.v3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YoutubeExcel
{
	public class YoutubeManager
	{
		// APIKey를 넣어주어야 한다.
		const string apiKey = "";

		YouTubeService service;

		public YoutubeManager()
		{
			service = new YouTubeService(new BaseClientService.Initializer()
			{
				ApiKey = apiKey,
				ApplicationName = "My YouTube Excel"
			});
		}

		public async Task<string> GetVideoTitle(string youtubeId)
		{
			var findRequest = service.Videos.List("snippet");
			findRequest.Id = youtubeId;
			findRequest.MaxResults = 25;

			var findResult = await findRequest.ExecuteAsync();

			var title = findResult.Items[0].Snippet.Title;

			string[] parsedTitle = title.Split('|');

			return parsedTitle[0];
		}

		public async Task<bool> GetVideoCaptionValid(string youtubeId)
		{
			var findRequest = service.Captions.List("snippet", youtubeId);

			var findResult = await findRequest.ExecuteAsync();

			var captionStatus = findResult.Items[0].Snippet.Status;

			if (captionStatus == "serving")
			{
				return true;
			}
			return false;
		}
	}
}
