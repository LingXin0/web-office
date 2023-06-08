<template>
  <div>{{ type }}</div>
</template>

<script>
import WebOfficeSDK from '../js/web-office-sdk-solution-v2.0.2.umd'

export default {
  name: "WebOffice",
  props: {
    type: String,
  },
  mounted() {
    const jssdk = WebOfficeSDK.init({
      officeType: 'w',
      appId: "SX20230607SAIWEY",
      fileId: "335d5e0e520c4ec68fc2c65ad269146c",
    });
    console.log(jssdk)

    // 需要等待 jssdk ready 之后再调用 API
  },
};
</script>
