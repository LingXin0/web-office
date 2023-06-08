<template>
    <div>{{ type }}</div>
</template>

<script>
import WebOfficeSDK from "../js/web-office-sdk-solution-v2.0.2.es.js";

export default {
    name: "WebOffice",
    props: {
        type: String,
    },
    data() {},
    mounted() {
        const instance = WebOfficeSDK.init({
            officeType: "w",
            appId: "SX20230607SAIWEY",
            fileId: "101e0bd54c8c4bc7976ee06941bc231d",
            token: "1",
            mount: "#weboffice",
            mode: "simple",
            wpsOptions: {
                isShowDocMap: false, // 是否开启目录功能，默认开启
                isBestScale: true, // 打开文档时，默认以最佳比例显示
            },
        });
        // 需要等待 jssdk ready 之后再调用 API
        instance.ready();
        console.log("instance", instance);
        console.log("aaa", JSON.parse(JSON.stringify(instance)));
        const app = instance.Application;
        console.log("app", app);

        // app.CommandBars("MoreMenus");

        // 书签对象
        const bookmarks = app.ActiveDocument.Bookmarks;

        // 添加书签
        bookmarks.Add({
            Name: "WebOffice",
            Range: {
                Start: 1,
                End: 10,
            },
        });

        // 1. 搜索并高亮文本
        const findResult = app.ActiveDocument.Find.Execute("WebOffice");

        // 2. 取消搜索结果高亮
        app.ActiveDocument.Find.ClearHitHighlight();

        // 3. 获取位置信息
        const { pos } = findResult[0];

        document
            .querySelector(".input1")
            .addEventListener("input", async (e) => {
                // 书签对象
                // const bookmarks = await app.ActiveDocument.Bookmarks;
                console.log("bookmarks:", bookmarks);
                // 替换书签内容
                const isReplaceSuccess = await bookmarks.ReplaceBookmark([
                    {
                        name: "WebOffice",
                        type: "text",
                        value: e.target.value,
                    },
                ]);
                console.log(isReplaceSuccess); // true
            });
    },
    methods: {
        async init() {},
    },
};
</script>
