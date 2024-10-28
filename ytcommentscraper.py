import googleapiclient.discovery
import xlsxwriter

# Set up your API key and YouTube API service
api_key = "AIzaSyCUXZ3v0KHj_9qj-m-9CUUPxMBka2kEboo"  # Replace with your API key
video_id = "HyIyd9joTTc"  # Replace with your video ID

# Initialize the YouTube API client
youtube = googleapiclient.discovery.build("youtube", "v3", developerKey=api_key)

# Define a function to get comments and their likes
def get_comments(youtube, video_id, max_comments=300):
    comments = []
    likes = []
    next_page_token = None

    while len(comments) < max_comments:
        request = youtube.commentThreads().list(
            part="snippet",
            videoId=video_id,
            maxResults=min(max_comments - len(comments), 100),  # Get up to 100 comments at a time
            pageToken=next_page_token
        )
        response = request.execute()

        for item in response['items']:
            comment = item['snippet']['topLevelComment']['snippet']['textDisplay']
            like_count = item['snippet']['topLevelComment']['snippet'].get('likeCount', 0)
            comments.append(comment)
            likes.append(like_count)
            if len(comments) >= max_comments:
                break

        next_page_token = response.get('nextPageToken')
        if not next_page_token:
            break

    return comments, likes

# Define a function to save comments and likes to an Excel file
def save_to_excel(comments, likes, filename='comments.xlsx'):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Comments")  # Write header for comments
    worksheet.write(0, 1, "Likes")      # Write header for likes

    for row, (comment, like) in enumerate(zip(comments, likes), start=1):
        worksheet.write(row, 0, comment)  # Write comments
        worksheet.write(row, 1, like)      # Write likes

    workbook.close()
    print(f"Comments and likes saved to {filename}")

# Call the function to get comments and likes
comments, likes = get_comments(youtube, video_id, max_comments=300)

# Print results
for i, (comment, like) in enumerate(zip(comments, likes), 1):
    print(f"{i}: {comment} (Likes: {like})")

# Save comments and likes to Excel
save_to_excel(comments, likes)


