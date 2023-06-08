<template>
    <div>{{ type }}</div>
    ∑∑∑∑
</template>

<script>
import webOfficeSdk from "../js/web-office-sdk-solution-v2.0.2.umd";

export default {
    name: "WebOffice",
    props: {
        type: String,
    },
    created() {
        console.log(123);
        const jssdk = webOfficeSdk.init({
            // 必填项
            officeType: webOfficeSdk.OfficeType.Writer,
            appId: "SX20230607SAIWEY",
            fileId: "335d5e0e520c4ec68fc2c65ad269146c",
        });

        // 需要等待 jssdk ready 之后再调用 API
        jssdk.ready();
    },
};
</script>
