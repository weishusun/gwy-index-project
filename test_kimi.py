import os
from openai import OpenAI

key = os.getenv("MOONSHOT_API_KEY")
print("repr(key) =", repr(key))
print("len(key)  =", len(key))

client = OpenAI(
    api_key=key,
    base_url="https://api.moonshot.cn/v1",
)

resp = client.models.list()
print("models count:", len(resp.data))
print("first model id:", resp.data[0].id)
