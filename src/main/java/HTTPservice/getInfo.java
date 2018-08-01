package HTTPservice;

import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;

import java.io.IOException;

public class getInfo {
    private OkHttpClient client = new OkHttpClient();
    public String run(String roomID) throws IOException {
        String url = "http://open.douyucdn.cn/api/RoomApi/room/" + roomID;
        Request request = new Request.Builder().url(url).build();
        Response response = client.newCall(request).execute();
        if (response.isSuccessful()) {
            String roomInfoJson;
            assert response.body() != null;
            roomInfoJson = response.body().string();
            return roomInfoJson;
        } else {
            throw new IOException("Unexpected code " + response);
        }
    }
}
