import * as React from "react";
import {
  CommandBar,
  ICommandBarItemProps,
} from "@fluentui/react/lib/CommandBar";
import { INewProps } from "./INewProps";
import { getTheme, concatStyleSets } from "@fluentui/react/lib/Styling";
import {
  FontWeights,
  IButtonStyles,
  IconButton,
  IContextualMenuItemStyles,
  IIconProps,
  mergeStyleSets,
  Modal,
} from "office-ui-fabric-react";
import Form from "./Form/Form";
import Upload from "./Upload/Upload";
import { useId, useBoolean } from "@fluentui/react-hooks";

const theme = getTheme();
const cancelIcon: IIconProps = { iconName: "Cancel" };
const itemStyles: Partial<IContextualMenuItemStyles> = {
  label: { fontSize: 18 },
  icon: { color: theme.palette.red },
  iconHovered: { color: theme.palette.redDark },
};
const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "stretch",
      fontWeight: FontWeights.semibold,
      padding: "12px 12px 14px 24px",
    },
  ],
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
  },
});
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: "auto",
    marginTop: "4px",
    marginRight: "2px",
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
import { Image, IImageProps, ImageFit } from "@fluentui/react/lib/Image";

// These props are defined up here so they can easily be applied to multiple Images.
// Normally specifying them inline would be fine.
const imageProps: IImageProps = {
  imageFit: ImageFit.contain,
  src: "https://previews.123rf.com/images/tribalium123/tribalium1231602/tribalium123160200024/52655719-website-under-construction-background-under-construction-template.jpg",
  // Show a border around the image (just for demonstration purposes)
  styles: (props) => ({
    root: { border: "1px solid " + props.theme.palette.neutralSecondary },
  }),
};

export default function New(props: INewProps) {
  const titleId = useId("title");
  //React Hooks to manipulate Show/Hide Modal Boolean
  const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] =
    useBoolean(false);
  const [isUploadModalOpen, { setTrue: show, setFalse: hide }] =
    useBoolean(false);

  //Command Bar Items
  const _items: ICommandBarItemProps[] = [
    {
      key: "newItem",
      text: "New",
      iconProps: { iconName: "Add" },
      onClick: showModal,
    },
    {
      key: "upload",
      text: "Upload",
      iconProps: { iconName: "Upload" },
      onClick: show,
    },
  ];

  /******************Building the Power App Url*********************************/
  //App Url:
  const appWebLink: string = `416c59ab-e270-4c92-9275-fb403aa3c3a1`;
  const appUrl: string = `https://apps.powerapps.com/play/${appWebLink}`;
  //Frame URL:
  const frameUrl: string = `${appUrl}`;
  return (
    <>
      <CommandBar
        items={_items}
        ariaLabel="Inbox actions"
        styles={itemStyles}
      />
      {
        //Modal
      }
      <div>
        <Modal
          titleAriaId={titleId}
          isOpen={isModalOpen}
          onDismiss={hideModal}
          isBlocking={false}
          containerClassName={contentStyles.container}
        >
          <div className={contentStyles.header}>
            <span id={titleId}>New File</span>
            <IconButton
              styles={iconButtonStyles}
              iconProps={cancelIcon}
              ariaLabel="Close popup modal"
              onClick={hideModal}
            />
          </div>
          <div className={contentStyles.body}>
            <Form
              description={""}
              context={props.context}
              siteUrl={props.context.pageContext.web.absoluteUrl}
            ></Form>

            {/* <iframe
            className={contentStyles.frame}
              src={frameUrl}
              frameBorder="0"
              scrolling="no"
              allow="geolocation *; microphone *; camera *; fullscreen *;"
              sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-forms allow-orientation-lock allow-downloads"
            ></iframe> */}
          </div>
        </Modal>
      </div>
      <div>
        <Modal
          titleAriaId={titleId}
          isOpen={isUploadModalOpen}
          onDismiss={show}
          isBlocking={false}
          containerClassName={contentStyles.container}
        >
          <div className={contentStyles.header}>
            <span id={titleId}>Upload new file</span>
            <IconButton
              styles={iconButtonStyles}
              iconProps={cancelIcon}
              ariaLabel="Close popup modal"
              onClick={hide}
            />
          </div>
          <div className={contentStyles.body}>
           {/*  <Image
              {...imageProps}
              alt='Example of the image fit value "contain" on an image wider than the frame.'
              width={400}
              height={350}
            /> */}
             <Upload
              description={""}
              context={props.context}
              siteUrl={props.context.pageContext.web.absoluteUrl}
            ></Upload>
          </div>
        </Modal>
      </div>
    </>
  );
}
